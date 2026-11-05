using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

[assembly: AssemblyTitle("Vefeyât Word Eklentisi")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: ComVisible(false)]

namespace Vefeyat
{
    // These dual interfaces are the Office COM contracts. No Office PIA, VSTO,
    // Office.js runtime or network service is required.
    [ComVisible(true), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, int connectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object addInInst, ref Array custom);
        [DispId(2)] void OnDisconnection(int removeMode, ref Array custom);
        [DispId(3)] void OnAddInsUpdate(ref Array custom);
        [DispId(4)] void OnStartupComplete(ref Array custom);
        [DispId(5)] void OnBeginShutdown(ref Array custom);
    }

    [ComVisible(true), Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)] [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }

    [ComVisible(true), Guid("2B0AF7B7-BD0D-4C70-9136-BED2DBD8F8F3")]
    [ProgId("Vefeyat.Connect"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private object application;
        private bool dialogOpen;

        public void OnConnection(object app, int connectMode, object addInInst, ref Array custom) { application = app; }
        public void OnDisconnection(int removeMode, ref Array custom) { application = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { application = null; }

        public string GetCustomUI(string ribbonId)
        {
            return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon><tabs><tab id=""VefeyatTab"" label=""Vefeyât"">
    <group id=""VefeyatDocument"" label=""Belge"">
      <button id=""VefeyatSelection"" label=""Seçili adı ara"" size=""large"" imageMso=""FindDialog""
        onAction=""SearchSelection"" screentip=""Seçili şahsın tarihini ekle""
        supertip=""Belgede şahıs adını seçin. Yerel kayıtlardan doğru şahsı seçerek vefat tarihini adın sonuna ekleyin."" />
      <button id=""VefeyatBrowse"" label=""Şahısları ara"" size=""large"" imageMso=""AddressBook""
        onAction=""BrowsePeople"" screentip=""Yerel şahıs veritabanını aç"" />
    </group>
    <group id=""VefeyatLocal"" label=""Verilerim"">
      <button id=""VefeyatFolder"" label=""Veri klasörü"" imageMso=""FileOpen"" onAction=""ShowDataFolder"" />
      <button id=""VefeyatHelp"" label=""Kullanım"" imageMso=""Help"" onAction=""ShowHelp"" />
    </group>
  </tab></tabs></ribbon>
</customUI>";
        }

        public void SearchSelection(object control) { OpenCatalog(true); }
        public void BrowsePeople(object control) { OpenCatalog(false); }

        public void ShowDataFolder(object control)
        {
            try
            {
                Directory.CreateDirectory(DataDirectory);
                System.Diagnostics.Process.Start("explorer.exe", "\"" + DataDirectory + "\"");
            }
            catch (Exception e) { Report(e); }
        }

        public void ShowHelp(object control)
        {
            MessageBox.Show("Belgede şahıs adını seçip Vefeyât > Seçili adı ara düğmesine basın.\n" +
                "Doğru şahsı ve tarih önizlemesini kontrol edip Tarihi ekle düğmesini kullanın.\n" +
                "Ad ve biçimi korunur; tarih sonuna eklenir. Ctrl+Z ile geri alabilirsiniz.\n\n" +
                "Şahısları ara: yerel arama, kayıt ekleme/düzenleme, CSV aktarımı ve yedekleme.\n" +
                "Arama ve ekleme tamamen çevrimdışıdır. Kaynak bağlantıları yalnız metin olarak saklanır.\n\n" +
                "Veri konumu: " + DataDirectory, "Vefeyât", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static string DataDirectory
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Vefeyat", "Data"); }
        }

        private void OpenCatalog(bool useSelection)
        {
            if (dialogOpen) return;
            object documentObject = null;
            object rangeObject = null;
            dialogOpen = true;
            try
            {
                if (application == null) throw new InvalidOperationException("Word bağlantısı hazır değil. Word’ü yeniden açın.");
                string query = "";
                string originalText = "";
                int storyType = 0;
                if (useSelection)
                {
                    dynamic word = application;
                    if (word.Documents.Count == 0) throw new InvalidOperationException("Önce bir Word belgesi açın ve şahıs adını seçin.");
                    documentObject = word.ActiveDocument;
                    dynamic selection = word.Selection;
                    if ((int)selection.Type != 2) throw new InvalidOperationException("Yalnız şahıs adını içeren bir metin seçin.");
                    rangeObject = selection.Range.Duplicate;
                    dynamic range = rangeObject;
                    originalText = (string)range.Text;
                    SelectionRules.ValidateName(originalText);
                    storyType = (int)range.StoryType;
                    query = originalText.Trim();
                }

                CatalogStore store = OpenStore();
                if (store == null) return;
                // Explicitly choosing the person in a modal dialog is the only
                // route to a document write. Browsing/editing never writes Word.
                Action<Person> insert = null;
                if (useSelection)
                {
                    object capturedDocument = documentObject;
                    object capturedRange = rangeObject;
                    string capturedText = originalText;
                    int capturedStory = storyType;
                    insert = delegate(Person person)
                    {
                        InsertDate(capturedDocument, capturedRange, capturedText, capturedStory, person);
                    };
                }
                using (var form = new CatalogForm(store, insert, query))
                {
                    dynamic word = application;
                    if ((int)word.Windows.Count > 0)
                        form.ShowDialog(new WordWindow(new IntPtr((int)word.ActiveWindow.Hwnd)));
                    else form.ShowDialog();
                }
            }
            catch (Exception e) { Report(e); }
            finally
            {
                Release(rangeObject);
                Release(documentObject);
                dialogOpen = false;
            }
        }

        private static CatalogStore OpenStore()
        {
            string seedPath = Path.Combine(Path.GetDirectoryName(typeof(Connect).Assembly.Location), "data", "people.csv");
            try { return new CatalogStore(DataDirectory, seedPath); }
            catch (InvalidDataException e)
            {
                if (MessageBox.Show(e.Message + "\n\nBir XML yedeğinden geri yüklemek ister misiniz?",
                    "Vefeyât: veriyi kurtar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2) != DialogResult.Yes) return null;
                using (var picker = new OpenFileDialog())
                {
                    picker.Title = "Geri yüklenecek Vefeyât XML yedeğini seçin";
                    picker.Filter = "Vefeyât yedeği (*.xml)|*.xml";
                    picker.InitialDirectory = DataDirectory;
                    if (picker.ShowDialog() != DialogResult.OK) return null;
                    CatalogStore.RestoreBackup(DataDirectory, picker.FileName);
                }
                return new CatalogStore(DataDirectory, seedPath);
            }
        }

        private void InsertDate(object capturedDocument, object capturedRange, string originalText, int storyType, Person person)
        {
            object selectionRangeObject = null;
            object tailObject = null;
            object insertObject = null;
            object undoObject = null;
            bool undoStarted = false;
            bool wrote = false;
            try
            {
                string date = SelectionRules.ValidateDate(person.DateText);
                dynamic word = application;
                dynamic doc = capturedDocument;
                dynamic range = capturedRange;
                if (word.Documents.Count == 0 || !SameComObject((object)word.ActiveDocument, capturedDocument))
                    throw new InvalidOperationException("Etkin belge değişti. Şahıs adını yeniden seçin.");
                if ((int)doc.ProtectionType != -1 || (bool)doc.ReadOnly)
                    throw new InvalidOperationException("Belge düzenlemeye kapalı. Düzenlenebilir bir belgede yeniden deneyin.");

                dynamic currentSelection = word.Selection;
                if ((int)currentSelection.Type != 2)
                    throw new InvalidOperationException("Word’deki seçim değişti. Şahıs adını yeniden seçin.");
                selectionRangeObject = currentSelection.Range;
                dynamic selected = selectionRangeObject;
                if (!SelectionRules.SameSelection(originalText, (string)range.Text, (string)selected.Text,
                    (int)range.Start, (int)range.End, (int)selected.Start, (int)selected.End,
                    storyType, (int)selected.StoryType) || !(bool)selected.IsEqual(range))
                    throw new InvalidOperationException("Word’deki seçim değişti. Şahıs adını yeniden seçin; tarih eklenmedi.");

                int trailing = SelectionRules.TrailingSpaces(originalText);
                int position = (int)range.End - trailing;
                tailObject = range.Duplicate;
                dynamic tail = tailObject;
                tail.SetRange(position, position);
                tail.MoveEnd(1, date.Length + trailing + 32); // wdCharacter; clamped by Word to story end.
                if (SelectionRules.HasDate(originalText, (string)tail.Text, date))
                    throw new InvalidOperationException("Bu tarih adın hemen sonrasında zaten var. Tekrar eklenmedi.");

                insertObject = range.Duplicate;
                dynamic insertion = insertObject;
                insertion.SetRange(position, position);
                // Word 2013+ supports a named, single-step undo record. Older
                // versions still retain Word's native undo for this insertion.
                try
                {
                    undoObject = word.UndoRecord;
                    ((dynamic)undoObject).StartCustomRecord("Vefeyât: tarih ekle");
                    undoStarted = true;
                }
                catch (COMException) { }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }

                insertion.InsertAfter(" " + date);
                wrote = true;
            }
            catch (Exception e)
            {
                if (wrote) throw new InvalidOperationException("Ekleme yapıldı. Belgeyi kontrol edin; gerekirse Ctrl+Z kullanın.", e);
                throw;
            }
            finally
            {
                if (undoStarted)
                {
                    try { ((dynamic)undoObject).EndCustomRecord(); } catch (Exception) { }
                }
                Release(undoObject);
                Release(insertObject);
                Release(tailObject);
                Release(selectionRangeObject);
            }
        }

        private static bool SameComObject(object first, object second)
        {
            if (first == null || second == null) return false;
            IntPtr a = IntPtr.Zero, b = IntPtr.Zero;
            try
            {
                a = Marshal.GetIUnknownForObject(first);
                b = Marshal.GetIUnknownForObject(second);
                return a == b;
            }
            finally
            {
                if (a != IntPtr.Zero) Marshal.Release(a);
                if (b != IntPtr.Zero) Marshal.Release(b);
            }
        }

        private static void Release(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj)) Marshal.ReleaseComObject(obj);
        }

        private static void Report(Exception e)
        {
            string message = e is COMException ? "Word işlemi tamamlanamadı. Belgeyi ve düzenleme iznini kontrol edin.\n" + e.Message : e.Message;
            MessageBox.Show(message, "Vefeyât", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private sealed class WordWindow : IWin32Window
        {
            private readonly IntPtr handle;
            public WordWindow(IntPtr value) { handle = value; }
            public IntPtr Handle { get { return handle; } }
        }
    }

}
