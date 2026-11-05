using System;
using System.IO;
using System.Linq;
using System.Text;
using Vefeyat;

internal static class CoreTests
{
    private static string root;
    private static int assertions;
    private const string Header = "Kısa Ad;Tam Ad;Orijinal Ad;Vefat Tarihi;Tanıtım;Kaynak Link\r\n";

    private static void Assert(bool condition, string message)
    {
        assertions++;
        if (!condition) throw new Exception("FAILED: " + message);
    }
    private static void Throws<T>(Action action, string message) where T : Exception
    {
        try { action(); }
        catch (T) { assertions++; return; }
        throw new Exception("FAILED (no " + typeof(T).Name + "): " + message);
    }
    private static string Csv(string name, string body)
    {
        string path = Path.Combine(root, name + ".csv");
        File.WriteAllText(path, Header + body, new UTF8Encoding(true));
        return path;
    }
    private static CatalogStore Store(string name, string seed)
    {
        return new CatalogStore(Path.Combine(root, name), seed);
    }

    private static void CsvRoundTrip()
    {
        string seed = Csv("roundtrip", "\"İbn Mâtürîdî\";\"Tam; Ad\";\"ابن\";\"(ö. 333/944)\";\"İlk satır\r\nİkinci \"\"alıntı\"\"\";\"https://example.invalid/a\"\r\n" +
            "İSMAİL;---;---;;Açıklama;\r\n");
        CatalogStore store = Store("roundtrip-data", seed);
        Assert(store.People.Count == 2, "two CSV records including multiline field");
        Person first = store.People[0];
        Assert(first.FullName == "Tam; Ad" && first.OriginalName == "ابن", "quoted semicolon and Arabic preserved");
        Assert(first.Description == "İlk satır\r\nİkinci \"alıntı\"", "multiline quoted description preserved");
        Assert(store.People[1].FullName == "" && store.People[1].OriginalName == "", "placeholder fields cleaned");
        string exported = Path.Combine(root, "export.csv");
        store.ExportCsv(exported);
        Assert(File.ReadAllBytes(exported).Take(3).SequenceEqual(new byte[] { 239, 187, 191 }), "Excel-compatible UTF-8 BOM");
        CatalogStore reimported = Store("roundtrip-new", exported);
        Assert(reimported.People[0].Description == first.Description && reimported.People[0].FullName == first.FullName, "CSV roundtrip");
        CatalogStore reopened = Store("roundtrip-data", null);
        Assert(reopened.People[0].Id == first.Id && reopened.People[0].Description == first.Description, "XML preserves ID and CRLF");
    }

    private static void SearchAndMutations()
    {
        CatalogStore store = Store("search", null);
        Person prefix = new Person { ShortName = "Mâtürîdî Şârihi", Description = "Âlim" };
        Person exact = new Person { ShortName = "Mâtürîdî", OriginalName = "ماتريدي", Description = "Türk kelâm âlimi" };
        Person ibn = new Person { ShortName = "İbn-i İsmâil", FullName = "İbni İsmail" };
        store.Save(prefix); store.Save(exact); store.Save(ibn);
        Assert(exact.Id.Length > 0, "Save assigns stable ID");
        Assert(store.Search("maturidi").First().Id == exact.Id, "exact name precedes prefix");
        Assert(store.Search("MATURIDI kelam alim").Count == 1 && store.Search("MATURIDI kelam alim")[0].Id == exact.Id, "accent-insensitive all-token search");
        Assert(store.Search("ISMAIL").Count == 1 && store.Search("ısmail").Count == 1, "Turkish dotted and dotless i");
        Assert(store.Search("ibni ismail").Count == 1, "ibni alias and punctuation handling");
        Assert(store.Search("ماتريدي").Count == 1, "Arabic search");
        Assert(store.Search("").Count == 3, "blank query browses catalog");
        Person detached = store.People[0]; detached.ShortName = "UNSAVED";
        Assert(store.People[0].ShortName != "UNSAVED", "People returns detached copies");
        Person searched = store.Search("maturidi")[0]; searched.ShortName = "UNSAVED";
        Assert(store.Search("maturidi")[0].ShortName != "UNSAVED", "Search returns detached copies");
        Throws<InvalidDataException>(() => store.Save(new Person { ShortName = "---" }), "blank placeholder name rejected");
        Throws<InvalidDataException>(() => store.Save(new Person { ShortName = "  " }), "blank name rejected");
        Person edit = store.People.First(p => p.Id == exact.Id); edit.DateText = "(ö. 333/944)"; store.Save(edit);
        Assert(Store("search", null).People.First(p => p.Id == exact.Id).DateText == edit.DateText, "edit persists");
        store.Delete(prefix);
        Assert(store.People.Count == 2 && !store.People.Any(p => p.Id == prefix.Id), "delete by exact ID");
    }

    private static void ImportTransactionsAndIdentity()
    {
        string original = Csv("import-original", "A;Tam;اصل;(ö. 1);Eski;https://example.invalid/a\r\n");
        CatalogStore store = Store("import", original);
        string id = store.People[0].Id;
        store.ImportCsv(original); store.ImportCsv(original);
        Assert(store.People.Count == 1 && store.People[0].Id == id, "repeated import idempotent");
        string update = Csv("import-update", "A;Tam;اصل;(ö. 2);Yeni;https://example.invalid/a\r\n");
        store.ImportCsv(update);
        Assert(store.People.Count == 1 && store.People[0].Id == id && store.People[0].DateText == "(ö. 2)", "unique identity updated preserving ID");
        string invalid = Csv("import-invalid", "B;;;;;\r\nC;only two\r\n");
        Throws<InvalidDataException>(() => store.ImportCsv(invalid), "invalid later CSV row rejected");
        Assert(store.People.Count == 1 && store.People[0].Id == id, "failed import makes no partial changes");
        string quotes = Csv("import-quotes", "\"not closed;;;;;\r\n");
        Throws<InvalidDataException>(() => store.ImportCsv(quotes), "unclosed quoted field rejected");
        string invalidUtf8 = Path.Combine(root, "invalid-utf8.csv");
        File.WriteAllBytes(invalidUtf8, new byte[] { 0xff, 0xfe, 0x41 });
        Throws<InvalidDataException>(() => store.ImportCsv(invalidUtf8), "non-UTF8 file rejected clearly");
        string ambiguous = Csv("ambiguous", "Same;;;(ö. 1);First;\r\nSame;;;(ö. 2);Second;\r\n");
        CatalogStore homonyms = Store("homonyms", ambiguous);
        Assert(homonyms.People.Count == 2 && homonyms.People.Select(p => p.Id).Distinct().Count() == 2, "homonyms preserved distinctly");
        homonyms.ImportCsv(ambiguous);
        Assert(homonyms.People.Count == 2, "homonym reimport idempotent");
        string uncertain = Csv("ambiguous-change", "Same;;;(ö. 3);Unknown match;\r\n");
        homonyms.ImportCsv(uncertain); homonyms.ImportCsv(uncertain);
        Assert(homonyms.People.Count == 3 && homonyms.People.Any(p => p.DateText == "(ö. 1)") && homonyms.People.Any(p => p.DateText == "(ö. 2)"), "ambiguous update never overwrites arbitrary person");
    }

    private static void ConcurrentWritesAndBackups()
    {
        CatalogStore a = Store("concurrent", null);
        CatalogStore b = Store("concurrent", null);
        Person first = new Person { ShortName = "First" };
        a.Save(first);
        Throws<InvalidOperationException>(() => b.Save(new Person { ShortName = "Lost update" }), "stale process write rejected");
        Assert(Store("concurrent", null).People.Count == 1, "concurrent save retains first process update");
        b.Reload(); b.Save(new Person { ShortName = "Second" });
        Assert(b.People.Count == 2, "reload permits fresh write");
        CatalogStore prior = Store("prior", null); prior.Restore(b.PreviousPath);
        Assert(prior.People.Count == 1 && prior.People[0].Id == first.Id, "previous version backup recoverable");
        string backup = Path.Combine(root, "backup.xml"); b.Backup(backup);
        b.Delete(b.People[0]); b.Restore(backup);
        Assert(b.People.Count == 2 && b.People.Any(p => p.Id == first.Id), "backup restore preserves original IDs");
        string dtd = Path.Combine(root, "dtd.xml");
        File.WriteAllText(dtd, "<!DOCTYPE vefeyat [<!ENTITY x SYSTEM 'file:///etc/passwd'>]><vefeyat version=\"1\">&x;</vefeyat>");
        Throws<System.Xml.XmlException>(() => b.Restore(dtd), "DTD/external entities rejected");
        Assert(b.People.Count == 2, "invalid restore leaves data intact");
        Throws<InvalidOperationException>(() => b.Backup(b.DataPath), "backup cannot overwrite live catalog");
        Throws<InvalidOperationException>(() => b.ExportCsv(b.PreviousPath), "export cannot overwrite recovery copy");
        Assert(!Directory.GetFiles(Path.GetDirectoryName(b.DataPath), ".catalog-*.tmp").Any(), "atomic save cleans temporary files");
        File.WriteAllText(b.DataPath, "broken XML");
        Throws<InvalidDataException>(() => Store("concurrent", null), "corruption does not silently reset data");
        CatalogStore.RestoreBackup(Path.GetDirectoryName(b.DataPath), backup);
        CatalogStore recovered = Store("concurrent", null);
        Assert(recovered.People.Count == 2, "standalone recovery works when constructor cannot open live XML");
        Assert(File.ReadAllText(recovered.PreviousPath) == "broken XML", "recovery retains damaged original for inspection");
    }

    public static int Main(string[] args)
    {
        root = Path.Combine(Path.GetTempPath(), "vefeyat-core-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            CsvRoundTrip(); SearchAndMutations(); ImportTransactionsAndIdentity(); ConcurrentWritesAndBackups();
            if (args.Length > 0)
            {
                CatalogStore real = Store("uploaded-csv", args[0]);
                Assert(real.People.Count == 7934, "uploaded CSV contains all 7934 records");
                real.ImportCsv(args[0]);
                Assert(real.People.Count == 7934, "uploaded CSV reimport does not duplicate records");
                Assert(real.Search("maturidi").Count > 0, "uploaded catalog diacritic-insensitive query");
                Console.WriteLine("Uploaded CSV: " + real.People.Count + " records; import and search verified.");
            }
            Console.WriteLine("PASS: " + assertions + " assertions.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally { try { Directory.Delete(root, true); } catch { } }
    }
}
