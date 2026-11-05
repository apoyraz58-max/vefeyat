using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Xml;

namespace Vefeyat
{
    public sealed class Person
    {
        public string Id = "";
        public string ShortName = "";
        public string FullName = "";
        public string OriginalName = "";
        public string DateText = "";
        public string Description = "";
        public string SourceUrl = "";

        public Person Clone() { return (Person)MemberwiseClone(); }
    }

    // All mutable state is stored in this user's data directory. No network API,
    // database service, Office PIA, or third-party assembly is used.
    public sealed class CatalogStore
    {
        private static readonly string[] Headers = { "Kısa Ad", "Tam Ad", "Orijinal Ad", "Vefat Tarihi", "Tanıtım", "Kaynak Link" };
        private readonly string directory;
        private List<Person> people = new List<Person>();
        private string revision;
        public string DataPath { get; private set; }
        public string PreviousPath { get; private set; }
        public List<Person> People { get { return Copy(people); } }

        private CatalogStore(string dataDirectory)
        {
            if (String.IsNullOrWhiteSpace(dataDirectory)) throw new ArgumentException("Veri klasörü belirtilmedi.");
            directory = Path.GetFullPath(dataDirectory);
            Directory.CreateDirectory(directory);
            DataPath = Path.Combine(directory, "catalog.xml");
            PreviousPath = Path.Combine(directory, "catalog.previous.xml");
        }

        public CatalogStore(string dataDirectory, string seedCsvPath) : this(dataDirectory)
        {
            using (AcquireLock())
            {
                if (File.Exists(DataPath)) Load();
                else
                {
                    List<Person> initial = String.IsNullOrEmpty(seedCsvPath) ? new List<Person>() : ReadCsv(seedCsvPath);
                    foreach (Person p in initial) p.Id = Guid.NewGuid().ToString("N");
                    Persist(initial);
                }
            }
        }

        public void Reload() { using (AcquireLock()) Load(); }

        public List<Person> Search(string query)
        {
            string term = Normalize(query);
            if (term.Length == 0) return Copy(people);
            string[] tokens = term.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return people.Select(p => new { Person = p, Short = Normalize(p.ShortName), Full = Normalize(p.FullName), Original = Normalize(p.OriginalName),
                    Text = Normalize(p.ShortName + " " + p.FullName + " " + p.OriginalName + " " + p.DateText + " " + p.Description) })
                .Where(x => tokens.All(t => x.Text.IndexOf(t, StringComparison.Ordinal) >= 0))
                .OrderBy(x => x.Short == term ? 0 : (x.Full == term || x.Original == term ? 1 : (x.Short.StartsWith(term, StringComparison.Ordinal) ? 2 :
                    (x.Short.IndexOf(term, StringComparison.Ordinal) >= 0 ? 3 : 4))))
                .ThenBy(x => x.Short, StringComparer.Ordinal).Select(x => x.Person.Clone()).ToList();
        }

        public static string Normalize(string value)
        {
            StringBuilder result = new StringBuilder();
            bool space = true;
            foreach (char c in (value ?? "").Normalize(NormalizationForm.FormKD))
            {
                UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(c);
                if (category == UnicodeCategory.NonSpacingMark || category == UnicodeCategory.SpacingCombiningMark || category == UnicodeCategory.EnclosingMark) continue;
                if (Char.IsWhiteSpace(c) || Char.IsPunctuation(c) || Char.IsSymbol(c))
                {
                    if (!space) result.Append(' ');
                    space = true;
                }
                else
                {
                    result.Append(c == '\u0131' ? 'i' : Char.ToLowerInvariant(c));
                    space = false;
                }
            }
            return String.Join(" ", result.ToString().Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(word => word == "ibni" ? "ibn" : word));
        }

        public void Save(Person person)
        {
            if (person == null) throw new ArgumentNullException("person");
            Person clean = CleanPerson(person);
            Validate(clean);
            using (AcquireLock())
            {
                EnsureCurrent();
                List<Person> next = Copy(people);
                if (String.IsNullOrEmpty(clean.Id))
                {
                    clean.Id = Guid.NewGuid().ToString("N");
                    next.Add(clean);
                }
                else
                {
                    int index = next.FindIndex(p => p.Id == clean.Id);
                    if (index < 0) throw new InvalidOperationException("Kayıt artık bulunamıyor. Listeyi yeniden açın.");
                    next[index] = clean;
                }
                Persist(next);
                person.Id = clean.Id;
            }
        }

        public void Delete(Person person)
        {
            if (person == null || String.IsNullOrEmpty(person.Id)) throw new ArgumentException("Silinecek kayıt seçilmedi.");
            using (AcquireLock())
            {
                EnsureCurrent();
                List<Person> next = Copy(people);
                if (next.RemoveAll(p => p.Id == person.Id) == 0) throw new InvalidOperationException("Kayıt artık bulunamıyor. Listeyi yeniden açın.");
                Persist(next);
            }
        }

        public void ImportCsv(string path)
        {
            // Parse and validate the complete file before acquiring the write lock.
            // An invalid final row cannot leave an earlier row partially imported.
            List<Person> incoming = ReadCsv(path);
            using (AcquireLock())
            {
                EnsureCurrent();
                List<Person> next = Copy(people);
                Dictionary<string, List<Person>> existing = next.GroupBy(Identity).ToDictionary(g => g.Key, g => g.ToList(), StringComparer.Ordinal);
                bool changed = false;
                foreach (IGrouping<string, Person> group in incoming.GroupBy(Identity))
                {
                    List<Person> available;
                    if (!existing.TryGetValue(group.Key, out available)) available = new List<Person>();
                    else available = new List<Person>(available);
                    List<Person> unmatched = new List<Person>();
                    foreach (Person imported in group)
                    {
                        int exact = available.FindIndex(p => SameContent(p, imported));
                        if (exact >= 0) available.RemoveAt(exact);
                        else unmatched.Add(imported);
                    }
                    if (available.Count == 1 && unmatched.Count == 1)
                    {
                        // A unique identity can be updated without changing its ID.
                        Person replacement = unmatched[0];
                        replacement.Id = available[0].Id;
                        next[next.FindIndex(p => p.Id == replacement.Id)] = replacement;
                        changed = true;
                    }
                    else
                    {
                        // Ambiguous homonyms are retained as distinct records. Never
                        // choose an arbitrary existing person to overwrite.
                        foreach (Person imported in unmatched)
                        {
                            imported.Id = Guid.NewGuid().ToString("N");
                            next.Add(imported);
                            changed = true;
                        }
                    }
                }
                if (changed) Persist(next);
            }
        }

        public void ExportCsv(string path)
        {
            GuardExternalPath(path);
            StringBuilder csv = new StringBuilder();
            AppendCsv(csv, Headers);
            foreach (Person p in people) AppendCsv(csv, Fields(p));
            AtomicExternalWrite(path, new UTF8Encoding(true).GetPreamble().Concat(new UTF8Encoding(false).GetBytes(csv.ToString())).ToArray());
        }

        public void Backup(string path)
        {
            GuardExternalPath(path);
            using (AcquireLock())
            {
                EnsureCurrent();
                AtomicExternalWrite(path, File.ReadAllBytes(DataPath));
            }
        }

        public void Restore(string path)
        {
            List<Person> restored = ReadXml(path);
            using (AcquireLock())
            {
                EnsureCurrent();
                Persist(restored);
            }
        }

        // Recovery remains available when a damaged live XML prevents the public
        // constructor from opening. Validate first; retain the replaced live file
        // as catalog.previous.xml, including when it was damaged.
        public static void RestoreBackup(string dataDirectory, string backupPath)
        {
            List<Person> restored = ReadXml(backupPath);
            CatalogStore store = new CatalogStore(dataDirectory);
            using (store.AcquireLock()) store.Persist(restored);
        }

        private static List<Person> Copy(IEnumerable<Person> source) { return source.Select(p => p.Clone()).ToList(); }
        private static string Clean(string text)
        {
            string value = (text ?? "").Trim();
            return value == "---" ? "" : value;
        }
        private static Person CleanPerson(Person p)
        {
            return new Person { Id = p.Id ?? "", ShortName = Clean(p.ShortName), FullName = Clean(p.FullName), OriginalName = Clean(p.OriginalName),
                DateText = Clean(p.DateText), Description = Clean(p.Description), SourceUrl = Clean(p.SourceUrl) };
        }
        private static void Validate(Person p)
        {
            if (String.IsNullOrWhiteSpace(p.ShortName)) throw new InvalidDataException("Kısa Ad alanı boş bırakılamaz.");
            foreach (string value in Fields(p)) XmlConvert.VerifyXmlChars(value ?? "");
            XmlConvert.VerifyXmlChars(p.Id ?? "");
        }
        private static string[] Fields(Person p) { return new string[] { p.ShortName, p.FullName, p.OriginalName, p.DateText, p.Description, p.SourceUrl }; }
        private static string Identity(Person p) { return String.Join("\u001f", new string[] { p.ShortName, p.FullName, p.OriginalName, p.SourceUrl }); }
        private static bool SameContent(Person left, Person right) { return Fields(left).SequenceEqual(Fields(right), StringComparer.Ordinal); }

        private static List<Person> ReadCsv(string path)
        {
            string csv;
            try { csv = new UTF8Encoding(false, true).GetString(File.ReadAllBytes(path)); }
            catch (DecoderFallbackException) { throw new InvalidDataException("CSV UTF-8 biçiminde olmalıdır. Excel'de CSV UTF-8 olarak kaydedip yeniden deneyin."); }
            csv = csv.TrimStart('\ufeff');
            List<string[]> rows = ParseCsv(csv);
            if (rows.Count == 0 || rows[0].Length != Headers.Length || !rows[0].Select(x => x.Trim()).SequenceEqual(Headers))
                throw new InvalidDataException("CSV başlıkları şu sırada olmalıdır: " + String.Join(";", Headers));
            List<Person> result = new List<Person>();
            for (int i = 1; i < rows.Count; i++)
            {
                string[] row = rows[i];
                if (row.All(String.IsNullOrWhiteSpace)) continue;
                if (row.Length != 6) throw new InvalidDataException("CSV " + (i + 1) + ". kayıtta altı sütun bulunmalıdır.");
                Person p = new Person { ShortName = Clean(row[0]), FullName = Clean(row[1]), OriginalName = Clean(row[2]), DateText = Clean(row[3]), Description = Clean(row[4]), SourceUrl = Clean(row[5]) };
                try { Validate(p); }
                catch (Exception error) { throw new InvalidDataException("CSV " + (i + 1) + ". kayıt geçersiz: " + error.Message, error); }
                result.Add(p);
            }
            return result;
        }

        private static List<string[]> ParseCsv(string text)
        {
            // Semicolon is the native format. Comma-separated exports are accepted
            // when the header contains no semicolons.
            int end = text.IndexOfAny(new char[] { '\r', '\n' });
            string header = end < 0 ? text : text.Substring(0, end);
            char delimiter = header.IndexOf(';') >= 0 ? ';' : ',';
            List<string[]> rows = new List<string[]>();
            List<string> fields = new List<string>();
            StringBuilder field = new StringBuilder();
            bool quoted = false, afterQuote = false, started = false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (quoted)
                {
                    if (c == '"')
                    {
                        if (i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                        else { quoted = false; afterQuote = true; }
                    }
                    else field.Append(c);
                    started = true;
                }
                else if (c == delimiter || c == '\r' || c == '\n')
                {
                    fields.Add(field.ToString()); field.Length = 0; afterQuote = false; started = false;
                    if (c != delimiter)
                    {
                        rows.Add(fields.ToArray()); fields.Clear();
                        if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    }
                }
                else if (c == '"' && !started && field.Length == 0) { quoted = true; started = true; }
                else
                {
                    if (afterQuote || c == '"') throw new InvalidDataException("CSV tırnak yapısı geçersiz.");
                    field.Append(c); started = true;
                }
            }
            if (quoted) throw new InvalidDataException("CSV'de kapanmamış tırnak var.");
            if (started || afterQuote || field.Length > 0 || fields.Count > 0) { fields.Add(field.ToString()); rows.Add(fields.ToArray()); }
            return rows;
        }

        private static void AppendCsv(StringBuilder target, IEnumerable<string> row)
        {
            target.Append(String.Join(";", row.Select(value => "\"" + (value ?? "").Replace("\"", "\"\"") + "\"")));
            target.Append("\r\n");
        }

        private IDisposable AcquireLock()
        {
            for (int attempt = 0; ; attempt++)
            {
                try { return new FileStream(Path.Combine(directory, "catalog.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
                catch (IOException)
                {
                    if (attempt >= 24) throw new IOException("Vefeyât verisi başka bir Word penceresinde kullanılıyor. Biraz sonra yeniden deneyin.");
                    Thread.Sleep(40);
                }
            }
        }
        private static string HashFile(string path)
        {
            if (!File.Exists(path)) return null;
            using (SHA256 hash = SHA256.Create())
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                return Convert.ToBase64String(hash.ComputeHash(file));
        }
        private void EnsureCurrent()
        {
            if (!String.Equals(revision, HashFile(DataPath), StringComparison.Ordinal))
                throw new InvalidOperationException("Veriler başka bir Word penceresinde değişti. Değişiklikleri korumak için işlem yapılmadı. Listeyi kapatıp yeniden açın.");
        }
        private void Load()
        {
            try { people = ReadXml(DataPath); revision = HashFile(DataPath); }
            catch (Exception error)
            {
                throw new InvalidDataException("Yerel veri dosyası okunamadı: " + DataPath + ". Önceki sürüm varsa " + PreviousPath + " dosyasını yedekleyin ve Yedeği Geri Yükle komutunu kullanın. " + error.Message, error);
            }
        }
        private static List<Person> ReadXml(string path)
        {
            XmlReaderSettings settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 64L * 1024L * 1024L };
            XmlDocument document = new XmlDocument { XmlResolver = null };
            using (XmlReader reader = XmlReader.Create(path, settings)) document.Load(reader);
            XmlElement root = document.DocumentElement;
            if (root == null || root.Name != "vefeyat" || root.GetAttribute("version") != "1") throw new InvalidDataException("Geçerli bir Vefeyât veri/yedek dosyası seçin.");
            List<Person> result = new List<Person>();
            HashSet<string> ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (XmlNode node in root.ChildNodes)
            {
                if (node.NodeType != XmlNodeType.Element) continue;
                if (node.Name != "person") throw new InvalidDataException("Yedekte tanınmayan bir kayıt var.");
                XmlElement item = (XmlElement)node;
                Person p = new Person { Id = item.GetAttribute("id"), ShortName = Text(item, "shortName"), FullName = Text(item, "fullName"),
                    OriginalName = Text(item, "originalName"), DateText = Text(item, "dateText"), Description = Text(item, "description"), SourceUrl = Text(item, "sourceUrl") };
                Validate(p);
                if (String.IsNullOrWhiteSpace(p.Id) || !ids.Add(p.Id)) throw new InvalidDataException("Yedekte boş veya yinelenen kayıt kimliği var.");
                result.Add(p);
            }
            return result;
        }
        private static string Text(XmlElement parent, string name)
        {
            XmlNode node = parent.SelectSingleNode(name);
            if (node == null) throw new InvalidDataException("Yedekte eksik alan var: " + name);
            return node.InnerText;
        }
        private void Persist(List<Person> next)
        {
            string temporary = Path.Combine(directory, ".catalog-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                XmlWriterSettings settings = new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = true, NewLineHandling = NewLineHandling.Entitize };
                using (FileStream stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    using (XmlWriter writer = XmlWriter.Create(stream, settings))
                    {
                        writer.WriteStartDocument(); writer.WriteStartElement("vefeyat"); writer.WriteAttributeString("version", "1");
                        foreach (Person p in next)
                        {
                            Validate(p);
                            writer.WriteStartElement("person"); writer.WriteAttributeString("id", p.Id);
                            writer.WriteElementString("shortName", p.ShortName); writer.WriteElementString("fullName", p.FullName);
                            writer.WriteElementString("originalName", p.OriginalName); writer.WriteElementString("dateText", p.DateText);
                            writer.WriteElementString("description", p.Description); writer.WriteElementString("sourceUrl", p.SourceUrl);
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement(); writer.WriteEndDocument(); writer.Flush();
                        stream.Flush(true);
                    }
                }
                if (File.Exists(DataPath)) File.Replace(temporary, DataPath, PreviousPath, true);
                else File.Move(temporary, DataPath);
                people = Copy(next); revision = HashFile(DataPath);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
        private void GuardExternalPath(string path)
        {
            string full = Path.GetFullPath(path);
            if (String.Equals(full, DataPath, StringComparison.OrdinalIgnoreCase) || String.Equals(full, PreviousPath, StringComparison.OrdinalIgnoreCase) ||
                String.Equals(full, Path.Combine(directory, "catalog.lock"), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Yedek veya dışa aktarım için çalışma dosyasından farklı bir konum seçin.");
        }
        private static void AtomicExternalWrite(string path, byte[] bytes)
        {
            string full = Path.GetFullPath(path);
            string temp = Path.Combine(Path.GetDirectoryName(full), ".vefeyat-" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                using (FileStream stream = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { stream.Write(bytes, 0, bytes.Length); stream.Flush(true); }
                if (File.Exists(full)) File.Replace(temp, full, null);
                else File.Move(temp, full);
            }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }
    }
}
