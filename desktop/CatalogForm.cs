using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Vefeyat
{
    // The catalogue UI is entirely local; source addresses are displayed as text.
    public sealed class CatalogForm : Form
    {
        private readonly CatalogStore store;
        private readonly Action<Person> insert;
        private readonly TextBox query = new TextBox();
        private readonly DataGridView results = new DataGridView();
        private readonly Label resultCount = new Label();
        private readonly TextBox[] detailFields = new TextBox[6];
        private readonly Button insertButton = new Button();
        private readonly Button editButton = new Button();
        private readonly Button deleteButton = new Button();
        private readonly Timer searchTimer = new Timer();
        private readonly Color accent = Color.FromArgb(35, 100, 76);
        private List<Person> visiblePeople = new List<Person>();

        public CatalogForm(CatalogStore store, Action<Person> insert, string initialQuery)
        {
            if (store == null) throw new ArgumentNullException("store");
            this.store = store;
            this.insert = insert;
            Text = "Vefeyât — Şahıslar";
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(1050, 710);
            MinimumSize = new Size(810, 680);
            Font = new Font("Segoe UI", 10F);
            BackColor = Color.FromArgb(247, 249, 247);
            AutoScaleMode = AutoScaleMode.Dpi;
            BuildLayout();
            searchTimer.Interval = 180;
            searchTimer.Tick += delegate { searchTimer.Stop(); RefreshResults(null); };
            query.TextChanged += delegate { searchTimer.Stop(); searchTimer.Start(); };
            query.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    searchTimer.Stop();
                    RefreshResults(null);
                    e.SuppressKeyPress = true;
                }
                if (e.KeyCode == Keys.Down && results.Rows.Count > 0)
                {
                    results.Focus();
                    e.SuppressKeyPress = true;
                }
            };
            query.Text = (initialQuery ?? "").Trim();
            searchTimer.Stop();
            RefreshResults(null);
            Shown += delegate { query.Focus(); query.SelectAll(); };
            FormClosed += delegate { searchTimer.Stop(); searchTimer.Dispose(); };
        }

        private void BuildLayout()
        {
            TableLayoutPanel page = new TableLayoutPanel();
            page.Dock = DockStyle.Fill;
            page.Padding = new Padding(20);
            page.ColumnCount = 1;
            page.RowCount = 5;
            page.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            page.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            page.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            page.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            page.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            Controls.Add(page);

            Label heading = new Label();
            heading.Text = "Vefeyât";
            heading.Font = new Font("Segoe UI", 23F, FontStyle.Bold);
            heading.ForeColor = accent;
            heading.Dock = DockStyle.Fill;
            page.Controls.Add(heading, 0, 0);

            TableLayoutPanel search = new TableLayoutPanel();
            search.Dock = DockStyle.Fill;
            search.ColumnCount = 3;
            search.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 56));
            search.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            search.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));
            Label searchLabel = new Label();
            searchLabel.Text = "&Ara";
            searchLabel.Dock = DockStyle.Fill;
            searchLabel.TextAlign = ContentAlignment.MiddleLeft;
            search.Controls.Add(searchLabel, 0, 0);
            query.Dock = DockStyle.Fill;
            query.AccessibleName = "Şahıs ara";
            query.Margin = new Padding(0, 7, 8, 0);
            search.Controls.Add(query, 1, 0);
            Button find = MakeButton("Ara", delegate { searchTimer.Stop(); RefreshResults(null); });
            find.Dock = DockStyle.Fill;
            search.Controls.Add(find, 2, 0);
            page.Controls.Add(search, 0, 1);

            TableLayoutPanel content = new TableLayoutPanel();
            content.Dock = DockStyle.Fill;
            content.ColumnCount = 2;
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 47));
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 53));
            content.RowCount = 1;
            content.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            page.Controls.Add(content, 0, 2);

            TableLayoutPanel left = new TableLayoutPanel();
            left.Dock = DockStyle.Fill;
            left.ColumnCount = 1;
            left.RowCount = 3;
            left.Padding = new Padding(0, 0, 12, 0);
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            left.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            left.RowStyles.Add(new RowStyle(SizeType.Absolute, 43));
            resultCount.Dock = DockStyle.Fill;
            resultCount.TextAlign = ContentAlignment.MiddleLeft;
            resultCount.ForeColor = Color.DimGray;
            left.Controls.Add(resultCount, 0, 0);

            results.Dock = DockStyle.Fill;
            results.BackgroundColor = Color.White;
            results.BorderStyle = BorderStyle.FixedSingle;
            results.ReadOnly = true;
            results.AllowUserToAddRows = false;
            results.AllowUserToDeleteRows = false;
            results.AllowUserToResizeRows = false;
            results.MultiSelect = false;
            results.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            results.RowHeadersVisible = false;
            results.AutoGenerateColumns = false;
            results.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            results.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            results.DefaultCellStyle.SelectionBackColor = accent;
            results.DefaultCellStyle.SelectionForeColor = Color.White;
            results.AccessibleName = "Arama sonuçları";
            DataGridViewTextBoxColumn name = new DataGridViewTextBoxColumn();
            name.HeaderText = "Şahıs";
            name.FillWeight = 60;
            name.SortMode = DataGridViewColumnSortMode.NotSortable;
            results.Columns.Add(name);
            DataGridViewTextBoxColumn date = new DataGridViewTextBoxColumn();
            date.HeaderText = "Vefat tarihi";
            date.FillWeight = 40;
            date.SortMode = DataGridViewColumnSortMode.NotSortable;
            results.Columns.Add(date);
            results.SelectionChanged += delegate { ShowSelected(); };
            left.Controls.Add(results, 0, 1);
            FlowLayoutPanel personActions = new FlowLayoutPanel();
            personActions.Dock = DockStyle.Fill;
            personActions.WrapContents = false;
            personActions.Controls.Add(MakeButton("&Yeni", AddPerson));
            SetupButton(editButton, "&Düzenle", EditPerson);
            personActions.Controls.Add(editButton);
            SetupButton(deleteButton, "&Sil", DeletePerson);
            personActions.Controls.Add(deleteButton);
            left.Controls.Add(personActions, 0, 2);
            content.Controls.Add(left, 0, 0);

            TableLayoutPanel right = new TableLayoutPanel();
            right.Dock = DockStyle.Fill;
            right.ColumnCount = 1;
            right.RowCount = 2;
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            right.RowStyles.Add(new RowStyle(SizeType.Absolute, 86));
            right.Controls.Add(BuildDetails(), 0, 0);
            TableLayoutPanel insertion = new TableLayoutPanel();
            insertion.Dock = DockStyle.Fill;
            insertion.RowCount = 2;
            insertion.ColumnCount = 1;
            insertion.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            insertion.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            SetupButton(insertButton, "Vefat tarihini belgeye ekle", InsertPerson);
            insertButton.BackColor = accent;
            insertButton.ForeColor = Color.White;
            insertButton.FlatStyle = FlatStyle.Flat;
            insertButton.Dock = DockStyle.Fill;
            insertion.Controls.Add(insertButton, 0, 0);
            Label instruction = new Label();
            instruction.Text = insert == null
                ? "Tarih eklemek için Word’de şahıs adını seçin ve Vefeyât sekmesindeki ekleme komutunu kullanın."
                : "Doğru şahsı seçin. Tarih, seçtiğiniz adın sonuna eklenir; ad korunur.";
            instruction.Dock = DockStyle.Fill;
            instruction.Font = new Font(Font.FontFamily, 9F);
            instruction.ForeColor = Color.DimGray;
            instruction.Padding = new Padding(0, 5, 0, 0);
            insertion.Controls.Add(instruction, 0, 1);
            right.Controls.Add(insertion, 0, 1);
            content.Controls.Add(right, 1, 0);

            FlowLayoutPanel dataActions = new FlowLayoutPanel();
            dataActions.Dock = DockStyle.Fill;
            dataActions.WrapContents = true;
            dataActions.Padding = new Padding(0, 5, 0, 0);
            dataActions.Controls.Add(MakeButton("CSV içe aktar", ImportCsv));
            dataActions.Controls.Add(MakeButton("CSV dışa aktar", ExportCsv));
            dataActions.Controls.Add(MakeButton("Yedek al", Backup));
            dataActions.Controls.Add(MakeButton("Yedeği geri yükle", Restore));
            page.Controls.Add(dataActions, 0, 3);

            TableLayoutPanel footer = new TableLayoutPanel();
            footer.Dock = DockStyle.Fill;
            footer.ColumnCount = 2;
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 86));
            Label local = new Label();
            local.Text = "Veriler bu bilgisayarda saklanır. İnternet bağlantısı gerekmez.";
            local.ForeColor = Color.DimGray;
            local.Dock = DockStyle.Fill;
            local.TextAlign = ContentAlignment.MiddleLeft;
            local.Font = new Font(Font.FontFamily, 9F);
            footer.Controls.Add(local, 0, 0);
            Button close = MakeButton("Kapat", delegate { Close(); });
            close.Dock = DockStyle.Fill;
            close.DialogResult = DialogResult.Cancel;
            CancelButton = close;
            footer.Controls.Add(close, 1, 0);
            page.Controls.Add(footer, 0, 4);
        }

        private Control BuildDetails()
        {
            TableLayoutPanel details = new TableLayoutPanel();
            details.Dock = DockStyle.Fill;
            details.ColumnCount = 1;
            details.RowCount = 12;
            string[] labels = { "Kısa ad", "Tam ad", "Orijinal ad", "Eklenecek vefat tarihi", "Tanıtım", "Kaynak adresi" };
            for (int i = 0; i < labels.Length; i++)
            {
                details.RowStyles.Add(new RowStyle(SizeType.Absolute, 23));
                details.RowStyles.Add(i == 4
                    ? new RowStyle(SizeType.Percent, 100)
                    : new RowStyle(SizeType.Absolute, i == 5 ? 43 : 30));
                Label label = new Label();
                label.Text = labels[i];
                label.Dock = DockStyle.Fill;
                label.ForeColor = accent;
                label.TextAlign = ContentAlignment.MiddleLeft;
                details.Controls.Add(label, 0, i * 2);
                TextBox value = new TextBox();
                value.Dock = DockStyle.Fill;
                value.ReadOnly = true;
                value.BackColor = Color.White;
                value.AccessibleName = labels[i];
                value.Margin = new Padding(0, 0, 0, 4);
                if (i == 4 || i == 5)
                {
                    value.Multiline = true;
                    value.ScrollBars = ScrollBars.Vertical;
                }
                if (i == 2) value.Font = new Font("Segoe UI", 12F);
                detailFields[i] = value;
                details.Controls.Add(value, 0, i * 2 + 1);
            }
            return details;
        }

        private Button MakeButton(string text, Action action)
        {
            Button button = new Button();
            SetupButton(button, text, action);
            return button;
        }

        private static void SetupButton(Button button, string text, Action action)
        {
            button.Text = text;
            button.AutoSize = true;
            button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            button.MinimumSize = new Size(64, 31);
            button.Margin = new Padding(0, 5, 7, 0);
            button.Padding = new Padding(6, 0, 6, 0);
            button.Click += delegate { action(); };
        }

        private Person SelectedPerson
        {
            get
            {
                if (results.CurrentRow == null) return null;
                int index = results.CurrentRow.Index;
                return index >= 0 && index < visiblePeople.Count ? visiblePeople[index] : null;
            }
        }

        private void RefreshResults(string selectedId)
        {
            try
            {
                List<Person> matches = store.Search(query.Text.Trim());
                visiblePeople = matches.Take(200).ToList();
                results.Rows.Clear();
                foreach (Person person in visiblePeople)
                    results.Rows.Add(person.ShortName, person.DateText);
                resultCount.Text = matches.Count == 0 ? "Kayıt bulunamadı."
                    : matches.Count > visiblePeople.Count
                        ? String.Format("{0:N0} kayıt • İlk {1} gösteriliyor; aramayı daraltın.", matches.Count, visiblePeople.Count)
                        : String.Format("{0:N0} kayıt", matches.Count);
                int selectedIndex = selectedId == null ? 0 : visiblePeople.FindIndex(p => p.Id == selectedId);
                if (selectedIndex >= 0 && selectedIndex < results.Rows.Count)
                    results.CurrentCell = results.Rows[selectedIndex].Cells[0];
                ShowSelected();
            }
            catch (Exception error) { ShowError(error); }
        }

        private void ShowSelected()
        {
            Person person = SelectedPerson;
            string[] values = person == null ? new string[6] : new[]
                { person.ShortName, person.FullName, person.OriginalName, person.DateText, person.Description, person.SourceUrl };
            for (int i = 0; i < detailFields.Length; i++)
                if (detailFields[i] != null) detailFields[i].Text = values[i] ?? "";
            editButton.Enabled = person != null;
            deleteButton.Enabled = person != null;
            insertButton.Enabled = insert != null && person != null && !String.IsNullOrWhiteSpace(person.DateText);
        }

        private void AddPerson()
        {
            using (PersonEditor editor = new PersonEditor(null))
            {
                if (editor.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    store.Save(editor.Person);
                    query.Text = editor.Person.ShortName;
                    searchTimer.Stop();
                    RefreshResults(editor.Person.Id);
                }
                catch (Exception error) { ShowError(error); }
            }
        }

        private void EditPerson()
        {
            Person person = SelectedPerson;
            if (person == null) return;
            using (PersonEditor editor = new PersonEditor(person))
            {
                if (editor.ShowDialog(this) != DialogResult.OK) return;
                try { store.Save(editor.Person); RefreshResults(editor.Person.Id); }
                catch (Exception error) { ShowError(error); }
            }
        }

        private void DeletePerson()
        {
            Person person = SelectedPerson;
            if (person == null) return;
            if (MessageBox.Show(this, "“" + person.ShortName + "” kaydı silinsin mi?", "Kaydı sil",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            try { store.Delete(person); RefreshResults(null); }
            catch (Exception error) { ShowError(error); }
        }

        private void InsertPerson()
        {
            Person person = SelectedPerson;
            if (insert == null || person == null || String.IsNullOrWhiteSpace(person.DateText)) return;
            try
            {
                insert(person);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception error) { ShowError(error); }
        }

        private void ImportCsv()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "Şahıs verilerini içe aktar";
                dialog.Filter = "CSV dosyaları (*.csv)|*.csv|Tüm dosyalar (*.*)|*.*";
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    store.ImportCsv(dialog.FileName);
                    RefreshResults(null);
                    MessageBox.Show(this, "CSV içe aktarıldı. Veriler bu bilgisayara kaydedildi.", "Vefeyât", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception error) { ShowError(error); }
            }
        }

        private void ExportCsv()
        {
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = "Tüm şahısları CSV olarak dışa aktar";
                dialog.Filter = "CSV dosyası (*.csv)|*.csv";
                dialog.FileName = "Vefeyat-sahislar.csv";
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try { store.ExportCsv(dialog.FileName); ShowSaved(dialog.FileName); }
                catch (Exception error) { ShowError(error); }
            }
        }

        private void Backup()
        {
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = "Yerel verilerin yedeğini kaydet";
                dialog.Filter = "Vefeyât yedeği (*.xml)|*.xml";
                dialog.FileName = "Vefeyat-yedek-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".xml";
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try { store.Backup(dialog.FileName); ShowSaved(dialog.FileName); }
                catch (Exception error) { ShowError(error); }
            }
        }

        private void Restore()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "Vefeyât yedeğini seçin";
                dialog.Filter = "Vefeyât yedeği (*.xml)|*.xml";
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                if (MessageBox.Show(this, "Seçilen yedek mevcut şahıs kayıtlarının yerine geçecek. Devam edilsin mi?",
                    "Yedeği geri yükle", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
                try
                {
                    store.Restore(dialog.FileName);
                    query.Clear();
                    searchTimer.Stop();
                    RefreshResults(null);
                    MessageBox.Show(this, "Yedek geri yüklendi.", "Vefeyât", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception error) { ShowError(error); }
            }
        }

        private void ShowSaved(string path)
        {
            MessageBox.Show(this, "Dosya kaydedildi:\n" + path, "Vefeyât", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowError(Exception error)
        {
            MessageBox.Show(this, error.Message, "Vefeyât — İşlem tamamlanamadı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private sealed class PersonEditor : Form
        {
            private readonly TextBox[] fields = new TextBox[6];
            public Person Person { get; private set; }

            public PersonEditor(Person source)
            {
                Text = source == null ? "Yeni şahıs" : "Şahıs kaydını düzenle";
                StartPosition = FormStartPosition.CenterParent;
                Size = new Size(610, 640);
                MinimumSize = new Size(490, 590);
                Font = new Font("Segoe UI", 10F);
                AutoScaleMode = AutoScaleMode.Dpi;
                Person = new Person();
                if (source != null) Person.Id = source.Id;
                TableLayoutPanel layout = new TableLayoutPanel();
                layout.Dock = DockStyle.Fill;
                layout.Padding = new Padding(18);
                layout.ColumnCount = 1;
                layout.RowCount = 14;
                Controls.Add(layout);
                string[] labels = { "Kısa ad *", "Tam ad", "Orijinal ad", "Vefat tarihi (belgeye aynen eklenecek ifade)", "Tanıtım", "Kaynak adresi" };
                string[] values = source == null ? new string[6] : new[]
                    { source.ShortName, source.FullName, source.OriginalName, source.DateText, source.Description, source.SourceUrl };
                for (int i = 0; i < fields.Length; i++)
                {
                    layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
                    layout.RowStyles.Add(i == 4 ? new RowStyle(SizeType.Percent, 100) : new RowStyle(SizeType.Absolute, 34));
                    Label label = new Label();
                    label.Text = labels[i];
                    label.Dock = DockStyle.Fill;
                    label.TextAlign = ContentAlignment.MiddleLeft;
                    layout.Controls.Add(label, 0, i * 2);
                    TextBox field = new TextBox();
                    field.Text = values[i] ?? "";
                    field.Dock = DockStyle.Fill;
                    field.AccessibleName = labels[i];
                    if (i == 4) { field.Multiline = true; field.AcceptsReturn = true; field.ScrollBars = ScrollBars.Vertical; }
                    fields[i] = field;
                    layout.Controls.Add(field, 0, i * 2 + 1);
                }
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
                Label hint = new Label();
                hint.Text = "Örnek tarih ifadesi: (ö. 680/1282). * Kısa ad zorunludur.";
                hint.ForeColor = Color.DimGray;
                hint.Dock = DockStyle.Fill;
                layout.Controls.Add(hint, 0, 12);
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 43));
                FlowLayoutPanel actions = new FlowLayoutPanel();
                actions.Dock = DockStyle.Fill;
                actions.FlowDirection = FlowDirection.RightToLeft;
                Button cancel = new Button();
                SetupButton(cancel, "Vazgeç", delegate { DialogResult = DialogResult.Cancel; Close(); });
                cancel.DialogResult = DialogResult.Cancel;
                CancelButton = cancel;
                actions.Controls.Add(cancel);
                Button save = new Button();
                SetupButton(save, "Kaydet", delegate
                {
                    if (String.IsNullOrWhiteSpace(fields[0].Text))
                    {
                        MessageBox.Show(this, "Kısa ad girin.", "Vefeyât", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        fields[0].Focus();
                        return;
                    }
                    Person.ShortName = fields[0].Text.Trim();
                    Person.FullName = fields[1].Text.Trim();
                    Person.OriginalName = fields[2].Text.Trim();
                    Person.DateText = fields[3].Text.Trim();
                    Person.Description = fields[4].Text.Trim();
                    Person.SourceUrl = fields[5].Text.Trim();
                    DialogResult = DialogResult.OK;
                    Close();
                });
                actions.Controls.Add(save);
                AcceptButton = save;
                layout.Controls.Add(actions, 0, 13);
            }
        }
    }
}
