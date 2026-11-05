using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Vefeyât Kurulum")]
[assembly: AssemblyVersion("1.1.0.0")]
[assembly: ComVisible(false)]

namespace Vefeyat.Setup
{
    internal static class InstallerProgram
    {
        [STAThread]
        private static void Main(string[] args)
        {
            if (Environment.OSVersion.Platform != PlatformID.Win32NT) return;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new InstallerForm(Array.IndexOf(args, "--uninstall") >= 0));
        }
    }

    internal sealed class InstallerForm : Form
    {
        private readonly RadioButton install = new RadioButton();
        private readonly RadioButton remove = new RadioButton();
        private readonly Button run = new Button();
        private readonly Label status = new Label();

        public InstallerForm(bool uninstall)
        {
            Text = "Vefeyât Kurulum";
            Font = new Font("Segoe UI", 10F);
            ClientSize = new Size(580, 340);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            var page = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(24), RowCount = 6, ColumnCount = 1 };
            page.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            page.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
            page.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            page.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            page.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            page.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            Controls.Add(page);
            page.Controls.Add(new Label { Text = "Vefeyât", Font = new Font(Font.FontFamily, 22F, FontStyle.Bold),
                ForeColor = Color.FromArgb(35, 100, 76), Dock = DockStyle.Fill }, 0, 0);
            page.Controls.Add(new Label { Text = "Word’ün bütün pencerelerini kapatın. Kurulum yalnız bu kullanıcı hesabına yapılır. Kişisel şahıs verileriniz korunur.", Dock = DockStyle.Fill }, 0, 1);
            install.Text = "Vefeyât’ı kur veya onar";
            remove.Text = "Vefeyât’ı kaldır";
            install.Dock = remove.Dock = DockStyle.Fill;
            install.Checked = !uninstall;
            remove.Checked = uninstall;
            page.Controls.Add(install, 0, 2);
            page.Controls.Add(remove, 0, 3);
            status.Dock = DockStyle.Fill;
            status.ForeColor = Color.DimGray;
            page.Controls.Add(status, 0, 4);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            var close = new Button { Text = "Kapat", AutoSize = true, Height = 34 };
            close.Click += delegate { Close(); };
            run.Text = uninstall ? "Kaldır" : "Kur";
            run.AutoSize = true;
            run.Height = 34;
            run.Click += delegate { Execute(); };
            install.CheckedChanged += delegate { run.Text = remove.Checked ? "Kaldır" : "Kur"; };
            remove.CheckedChanged += delegate { run.Text = remove.Checked ? "Kaldır" : "Kur"; };
            actions.Controls.Add(close);
            actions.Controls.Add(run);
            page.Controls.Add(actions, 0, 5);
            AcceptButton = run;
            CancelButton = close;
        }

        private void Execute()
        {
            bool uninstall = remove.Checked;
            if (uninstall && MessageBox.Show(this, "Vefeyât Word eklentisi kaldırılsın mı? Şahıs kayıtlarınız korunacaktır.",
                "Vefeyât", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            run.Enabled = false;
            UseWaitCursor = true;
            status.Text = "Dosyalar ve kurulum durumu kontrol ediliyor…";
            status.Refresh();
            try
            {
                using (var gate = new Mutex(false, "Local\\Vefeyat.UserInstaller"))
                {
                    bool acquired;
                    try { acquired = gate.WaitOne(0); }
                    catch (AbandonedMutexException) { acquired = true; }
                    if (!acquired) throw new InvalidOperationException("Başka bir Vefeyât kurulumu açık. Onu kapatıp yeniden deneyin.");
                    try { if (uninstall) InstallerEngine.Uninstall(); else InstallerEngine.Install(); }
                    finally { gate.ReleaseMutex(); }
                }
                status.Text = uninstall ? "Eklenti kaldırıldı. Kişisel verileriniz korundu. Kurulum programının kopyası diskte bırakıldı."
                    : "Kurulum tamamlandı. Word’ü açıp Vefeyât sekmesini kullanabilirsiniz.";
                MessageBox.Show(this, status.Text, "Vefeyât", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception error)
            {
                status.Text = "İşlem tamamlanamadı.";
                MessageBox.Show(this, error.Message, "Vefeyât", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { run.Enabled = true; UseWaitCursor = false; }
        }
    }

    internal static class InstallerEngine
    {
        internal const string ProgId = "Vefeyat.Connect";
        internal const string ClassId = "{2B0AF7B7-BD0D-4C70-9136-BED2DBD8F8F3}";
        internal const string AssemblyFile = "Vefeyat.WordAddin.dll";
        internal const string InstallerFile = "Vefeyat.Kurulum.exe";
        internal const string UninstallPath = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\Vefeyat";
        private static readonly string ClassPath = @"Software\Classes\CLSID\" + ClassId;
        private static readonly string ProgIdPath = @"Software\Classes\" + ProgId;
        private static readonly string AddinPath = @"Software\Microsoft\Office\Word\Addins\" + ProgId;
        private static string[] OwnedRegistryRoots { get { return new[] { ClassPath, ProgIdPath, AddinPath, UninstallPath }; } }
        internal static string InstallDirectory
        {
            get
            {
                string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (String.IsNullOrWhiteSpace(local) || !Path.IsPathRooted(local))
                    throw new InvalidOperationException("Windows yerel kullanıcı veri klasörü bulunamadı.");
                return Path.Combine(local, "Vefeyat", "Addin");
            }
        }

        internal static Dictionary<string, Dictionary<string, object>> RegistrationPlan(AssemblyName assembly,
            string libraryPath, string installerPath, string publisher)
        {
            if (assembly == null || assembly.Name != "Vefeyat.WordAddin" || assembly.Version == null)
                throw new InvalidDataException("Eklenti dosyasının kimliği geçersiz.");
            libraryPath = Path.GetFullPath(libraryPath);
            installerPath = Path.GetFullPath(installerPath);
            var registration = new Dictionary<string, object> { { "", "mscoree.dll" }, { "ThreadingModel", "Both" },
                { "Class", ProgId }, { "Assembly", assembly.FullName }, { "RuntimeVersion", "v4.0.30319" },
                { "CodeBase", new Uri(libraryPath).AbsoluteUri } };
            return new Dictionary<string, Dictionary<string, object>>
            {
                { ClassPath, new Dictionary<string, object> { { "", ProgId } } },
                { ClassPath + @"\InprocServer32", registration },
                { ClassPath + @"\InprocServer32\" + assembly.Version, new Dictionary<string, object>(registration) },
                { ClassPath + @"\ProgId", new Dictionary<string, object> { { "", ProgId } } },
                { ProgIdPath, new Dictionary<string, object> { { "", "Vefeyât Word Eklentisi" } } },
                { ProgIdPath + @"\CLSID", new Dictionary<string, object> { { "", ClassId } } },
                { AddinPath, new Dictionary<string, object> { { "FriendlyName", "Vefeyât" }, { "Description", "Bilgisayarınızdaki şahıs kataloğu" },
                    { "LoadBehavior", 3 }, { "CommandLineSafe", 0 } } },
                { UninstallPath, new Dictionary<string, object> { { "DisplayName", "Vefeyât Word Eklentisi" },
                    { "DisplayVersion", assembly.Version.ToString() }, { "Publisher", publisher ?? "" },
                    { "InstallLocation", Path.GetDirectoryName(libraryPath) }, { "DisplayIcon", installerPath },
                    { "UninstallString", Quote(installerPath) + " --uninstall" }, { "NoModify", 1 }, { "NoRepair", 1 } } }
            };
        }

        internal static void Install()
        {
            EnsureWordClosed();
            EnsureWordInstalled();
            string source = Path.GetDirectoryName(typeof(InstallerProgram).Assembly.Location);
            string sourceExe = typeof(InstallerProgram).Assembly.Location;
            string sourceLibrary = Path.Combine(source, AssemblyFile);
            string sourceSeed = Path.Combine(source, "data", "people.csv");
            // Hold read-only handles across signature checks and copying. A package
            // file cannot be swapped between verification and installation.
            using (var executable = OpenPayload(sourceExe))
            using (var library = OpenPayload(sourceLibrary))
            using (var seed = OpenPayload(sourceSeed))
            {
                SignedPublisher publisher = AuthenticodeVerifier.Verify(sourceExe, executable);
                SignedPublisher libraryPublisher = AuthenticodeVerifier.Verify(sourceLibrary, library);
                if (!String.Equals(publisher.Thumbprint, libraryPublisher.Thumbprint, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("Kurulum programı ile eklentinin yayıncısı eşleşmiyor. Doğru dağıtım paketini kullanın.");
                ValidateSeed(seed);
                AssemblyName name = AssemblyName.GetAssemblyName(sourceLibrary);
                string installedLibrary = Path.Combine(InstallDirectory, AssemblyFile);
                string installedExe = Path.Combine(InstallDirectory, InstallerFile);
                var plan = RegistrationPlan(name, installedLibrary, installedExe, publisher.Subject);
                ValidateOwnership();
                var registryBackups = CaptureRegistry();
                var backups = new List<FileBackup>();
                bool registryChanged = false;
                try
                {
                    EnsureWordClosed();
                    CopyPayload(library, sourceLibrary, installedLibrary, backups);
                    CopyPayload(seed, sourceSeed, Path.Combine(InstallDirectory, "data", "people.csv"), backups);
                    CopyPayload(executable, sourceExe, installedExe, backups);
                    EnsureWordClosed();
                    registryChanged = true;
                    foreach (RegistryView view in Views())
                    using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                    {
                        root.DeleteSubKeyTree(ClassPath, false); // Remove obsolete owned COM versions.
                        foreach (var entry in plan)
                        using (RegistryKey key = root.CreateSubKey(entry.Key))
                            foreach (var value in entry.Value)
                                key.SetValue(value.Key, value.Value, value.Value is int ? RegistryValueKind.DWord : RegistryValueKind.String);
                    }
                }
                catch (Exception error) { throw Rollback(error, registryChanged ? registryBackups : null, backups); }
            }
        }

        internal static void Uninstall()
        {
            EnsureWordClosed();
            string sourceExe = typeof(InstallerProgram).Assembly.Location;
            using (var executable = OpenPayload(sourceExe)) AuthenticodeVerifier.Verify(sourceExe, executable);
            ValidateOwnership();
            var registryBackups = CaptureRegistry();
            var backups = new List<FileBackup>();
            bool registryChanged = false;
            try
            {
                EnsureWordClosed();
                // Keep the running signed installer. Windows cannot remove a
                // running executable; no delayed shell/self-delete is scheduled.
                foreach (string relative in new[] { AssemblyFile, Path.Combine("data", "people.csv") })
                {
                    string path = Path.Combine(InstallDirectory, relative);
                    if (File.Exists(path)) { var backup = FileBackup.Capture(path); backups.Add(backup); File.Delete(path); }
                }
                EnsureWordClosed();
                registryChanged = true;
                foreach (RegistryView view in Views())
                using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                    foreach (string path in OwnedRegistryRoots) root.DeleteSubKeyTree(path, false);
            }
            catch (Exception error) { throw Rollback(error, registryChanged ? registryBackups : null, backups); }
        }

        private static string Quote(string path) { return "\"" + path + "\""; }
        private static IEnumerable<RegistryView> Views()
        {
            yield return RegistryView.Registry32;
            if (Environment.Is64BitOperatingSystem) yield return RegistryView.Registry64;
        }
        private static void EnsureWordClosed()
        {
            Process[] processes = Process.GetProcessesByName("WINWORD");
            bool found = processes.Length > 0;
            foreach (Process process in processes) process.Dispose();
            if (found) throw new InvalidOperationException("Önce bütün Word pencerelerini kapatın; ardından yeniden deneyin.");
        }
        private static void EnsureWordInstalled()
        {
            foreach (RegistryView view in Views())
            using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, view))
            using (RegistryKey key = root.OpenSubKey(@"Word.Application\CLSID"))
                if (key != null && !String.IsNullOrWhiteSpace(key.GetValue("") as string)) return;
            throw new InvalidOperationException("Bu kullanıcı için Windows masaüstü Microsoft Word kurulu olmalıdır.");
        }
        private static FileStream OpenPayload(string path)
        {
            if (!File.Exists(path)) throw new FileNotFoundException("Kurulum paketi eksik. ZIP’in bütün dosyalarını aynı klasöre çıkarın.", Path.GetFileName(path));
            var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            if (stream.Length > 64L * 1024L * 1024L) { stream.Dispose(); throw new InvalidDataException("Kurulum dosyası beklenenden büyük."); }
            return stream;
        }
        private static void ValidateSeed(FileStream seed)
        {
            seed.Position = 0;
            using (var reader = new StreamReader(seed, new UTF8Encoding(false, true), true, 4096, true))
            {
                string header = reader.ReadLine();
                if (header == null || header.Replace("\"", "") != "Kısa Ad;Tam Ad;Orijinal Ad;Vefat Tarihi;Tanıtım;Kaynak Link")
                    throw new InvalidDataException("Başlangıç veri dosyasının başlıkları geçersiz.");
            }
            seed.Position = 0;
        }
        private static void ValidateOwnership()
        {
            foreach (RegistryView view in Views())
            using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
            {
                using (RegistryKey key = root.OpenSubKey(ClassPath + @"\InprocServer32"))
                {
                    if (key != null)
                    {
                        if (!String.Equals(key.GetValue("Class", "") as string, ProgId, StringComparison.Ordinal))
                            throw new InvalidOperationException("Eklenti kimliği başka bir uygulamaya ait; kurulum değiştirilmedi.");
                        string codebase = key.GetValue("CodeBase", "") as string;
                        Uri uri;
                        if (!String.IsNullOrEmpty(codebase) && (!Uri.TryCreate(codebase, UriKind.Absolute, out uri) || !uri.IsFile ||
                            !String.Equals(Path.GetFullPath(uri.LocalPath), Path.Combine(InstallDirectory, AssemblyFile), StringComparison.OrdinalIgnoreCase)))
                            throw new InvalidOperationException("Eklenti başka bir konuma kayıtlı. Mevcut kurulumun yöneticisiyle görüşün.");
                    }
                }
                using (RegistryKey key = root.OpenSubKey(ProgIdPath + @"\CLSID"))
                    if (key != null && !String.Equals(key.GetValue("", "") as string, ClassId, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("Program kimliği başka bir uygulamaya bağlı; işlem yapılmadı.");
                using (RegistryKey key = root.OpenSubKey(UninstallPath))
                    if (key != null && !String.Equals(key.GetValue("InstallLocation", "") as string, InstallDirectory, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("Kaldırma kaydı başka bir kurulum konumuna ait; işlem yapılmadı.");
            }
            string installed = Path.Combine(InstallDirectory, AssemblyFile);
            if (File.Exists(installed) && AssemblyName.GetAssemblyName(installed).Name != "Vefeyat.WordAddin")
                throw new InvalidOperationException("Hedef klasörde başka bir uygulamanın dosyası var; işlem yapılmadı.");
        }

        private static void CopyPayload(FileStream input, string source, string destination, List<FileBackup> backups)
        {
            if (String.Equals(Path.GetFullPath(source), Path.GetFullPath(destination), StringComparison.OrdinalIgnoreCase)) return;
            FileBackup backup = FileBackup.Capture(destination);
            backups.Add(backup);
            Directory.CreateDirectory(Path.GetDirectoryName(destination));
            string temporary = destination + ".new-" + Guid.NewGuid().ToString("N");
            try
            {
                input.Position = 0;
                using (var output = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                { input.CopyTo(output); output.Flush(true); }
                if (File.Exists(destination)) File.Replace(temporary, destination, null);
                else File.Move(temporary, destination);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private static List<RegistryBackup> CaptureRegistry()
        {
            var backups = new List<RegistryBackup>();
            foreach (RegistryView view in Views())
            using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                foreach (string path in OwnedRegistryRoots)
                using (RegistryKey key = root.OpenSubKey(path))
                    backups.Add(new RegistryBackup { View = view, Path = path, Tree = key == null ? null : RegistryTree.Capture(key) });
            return backups;
        }
        private static Exception Rollback(Exception error, List<RegistryBackup> registry, List<FileBackup> files)
        {
            var failures = new List<string>();
            if (registry != null)
                for (int i = registry.Count - 1; i >= 0; i--)
                    try { registry[i].Restore(); } catch (Exception e) { failures.Add(e.Message); }
            for (int i = files.Count - 1; i >= 0; i--)
                try { files[i].Restore(); } catch (Exception e) { failures.Add(e.Message); }
            return failures.Count == 0 ? new InvalidOperationException(error.Message + "\nÖnceki kurulum durumu korundu.", error)
                : new InvalidOperationException(error.Message + "\nÖnceki kurulum geri yüklenirken sorun oluştu: " + String.Join("; ", failures), error);
        }

        private sealed class FileBackup
        {
            public string Path;
            public byte[] Bytes;
            public static FileBackup Capture(string path)
            {
                if (File.Exists(path) && new FileInfo(path).Length > 64L * 1024L * 1024L) throw new IOException("Mevcut kurulum dosyası beklenenden büyük.");
                return new FileBackup { Path = path, Bytes = File.Exists(path) ? File.ReadAllBytes(path) : null };
            }
            public void Restore()
            {
                if (Bytes == null) { if (File.Exists(Path)) File.Delete(Path); return; }
                string temp = Path + ".restore-" + Guid.NewGuid().ToString("N");
                try
                {
                    File.WriteAllBytes(temp, Bytes);
                    if (File.Exists(Path)) File.Replace(temp, Path, null); else File.Move(temp, Path);
                }
                finally { if (File.Exists(temp)) File.Delete(temp); }
            }
        }
        private sealed class RegistryBackup
        {
            public RegistryView View;
            public string Path;
            public RegistryTree Tree;
            public void Restore()
            {
                using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, View))
                {
                    root.DeleteSubKeyTree(Path, false);
                    if (Tree != null) using (RegistryKey key = root.CreateSubKey(Path)) Tree.Restore(key);
                }
            }
        }
        private sealed class RegistryTree
        {
            private readonly Dictionary<string, object> values = new Dictionary<string, object>();
            private readonly Dictionary<string, RegistryValueKind> kinds = new Dictionary<string, RegistryValueKind>();
            private readonly Dictionary<string, RegistryTree> children = new Dictionary<string, RegistryTree>();
            public static RegistryTree Capture(RegistryKey key)
            {
                var tree = new RegistryTree();
                foreach (string name in key.GetValueNames())
                { tree.values[name] = key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames); tree.kinds[name] = key.GetValueKind(name); }
                foreach (string name in key.GetSubKeyNames()) using (RegistryKey child = key.OpenSubKey(name)) tree.children[name] = Capture(child);
                return tree;
            }
            public void Restore(RegistryKey key)
            {
                foreach (var value in values) key.SetValue(value.Key, value.Value, kinds[value.Key]);
                foreach (var child in children) using (RegistryKey subkey = key.CreateSubKey(child.Key)) child.Value.Restore(subkey);
            }
        }
    }

    internal sealed class SignedPublisher
    {
        public string Thumbprint;
        public string Subject;
    }

    internal static class AuthenticodeVerifier
    {
        // Windows trust validation is a prerequisite, not a guarantee that Smart
        // App Control or a managed device policy accepts a release. Cache-only
        // validation never downloads trust material. The local verifier checks
        // signature integrity, the Windows trust chain and validity. This local
        // offline check does not perform revocation checks (including cached
        // revocation checks) and does not establish Microsoft public-root/SAC
        // approval. Releases require a public-trust code-signing provider;
        // Windows/SAC policy remains independent and is never changed here.
        internal static SignedPublisher Verify(string path, FileStream heldFile)
        {
            var file = new WinTrustFileInfo { Size = (uint)Marshal.SizeOf(typeof(WinTrustFileInfo)), Path = Path.GetFullPath(path),
                File = heldFile.SafeFileHandle.DangerousGetHandle(), KnownSubject = IntPtr.Zero };
            IntPtr filePointer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WinTrustFileInfo)));
            Marshal.StructureToPtr(file, filePointer, false);
            var trust = new WinTrustData { Size = (uint)Marshal.SizeOf(typeof(WinTrustData)), UIChoice = 2, UnionChoice = 1,
                FileInfo = filePointer, StateAction = 1, ProviderFlags = 0x1000 | 0x10 | 0x2000 };
            Guid action = new Guid("00AAC56B-CD44-11D0-8CC2-00C04FC295EE");
            try
            {
                int result = WinVerifyTrust(IntPtr.Zero, ref action, ref trust);
                if (result != 0)
                    throw new InvalidDataException("Dijital imza bu bilgisayarın çevrimdışı güven verileriyle doğrulanamadı: " +
                        Path.GetFileName(path) + ". Yayıncıdan geçerli imzalı paketi isteyin. Kod: 0x" + unchecked((uint)result).ToString("X8"));
                IntPtr provider = WTHelperProvDataFromStateData(trust.StateData);
                IntPtr signerPointer = provider == IntPtr.Zero ? IntPtr.Zero : WTHelperGetProvSignerFromChain(provider, 0, false, 0);
                if (signerPointer == IntPtr.Zero) throw new InvalidDataException("Doğrulanmış yayıncı bilgisi okunamadı.");
                var signer = (ProviderSigner)Marshal.PtrToStructure(signerPointer, typeof(ProviderSigner));
                if (signer.CertificateCount == 0 || signer.Certificates == IntPtr.Zero) throw new InvalidDataException("İmzanın yayıncı sertifikası bulunamadı.");
                var first = (ProviderCertificate)Marshal.PtrToStructure(signer.Certificates, typeof(ProviderCertificate));
                if (first.Certificate == IntPtr.Zero) throw new InvalidDataException("Yayıncı sertifikası okunamadı.");
                using (var certificate = new X509Certificate2(first.Certificate))
                    return new SignedPublisher { Thumbprint = certificate.Thumbprint, Subject = certificate.Subject };
            }
            finally
            {
                try { trust.StateAction = 2; WinVerifyTrust(IntPtr.Zero, ref action, ref trust); }
                finally { Marshal.DestroyStructure(filePointer, typeof(WinTrustFileInfo)); Marshal.FreeHGlobal(filePointer); }
            }
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WinTrustFileInfo
        {
            public uint Size;
            [MarshalAs(UnmanagedType.LPWStr)] public string Path;
            public IntPtr File;
            public IntPtr KnownSubject;
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WinTrustData
        {
            public uint Size;
            public IntPtr PolicyCallback;
            public IntPtr SIPClient;
            public uint UIChoice;
            public uint RevocationChecks;
            public uint UnionChoice;
            public IntPtr FileInfo;
            public uint StateAction;
            public IntPtr StateData;
            public IntPtr URLReference;
            public uint ProviderFlags;
            public uint UIContext;
            public IntPtr SignatureSettings;
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct ProviderSigner
        {
            public uint Size;
            public System.Runtime.InteropServices.ComTypes.FILETIME VerifyAsOf;
            public uint CertificateCount;
            public IntPtr Certificates;
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct ProviderCertificate
        {
            public uint Size;
            public IntPtr Certificate;
        }
        [DllImport("wintrust.dll", ExactSpelling = true, CharSet = CharSet.Unicode)]
        private static extern int WinVerifyTrust(IntPtr window, ref Guid action, ref WinTrustData data);
        [DllImport("wintrust.dll", ExactSpelling = true)]
        private static extern IntPtr WTHelperProvDataFromStateData(IntPtr stateData);
        [DllImport("wintrust.dll", ExactSpelling = true)]
        private static extern IntPtr WTHelperGetProvSignerFromChain(IntPtr providerData, uint signerIndex,
            [MarshalAs(UnmanagedType.Bool)] bool counterSigner, uint counterSignerIndex);
    }
}
