using System.Data;

namespace AccountTester
{
    public partial class OptionPrinter : Form
    {
        static string T(string key) => LangManager.Instance.Translate(key);

        public OptionPrinter()
        {
            InitializeComponent();

            UpdateTexts();
            LangManager.Instance.LanguageChanged += UpdateTexts;

            string[] printerList = Variables.PrinterList.Split(';').Select(p => p.Trim()).Where(p => !string.IsNullOrEmpty(p)).ToArray();

            foreach (var printer in printerList)
            {
                checkedListBoxPrinter.Items.Add(printer, true);
            }
        }

        private void UpdateTexts()
        {
            this.Text = T("PrinterOptions");
            labelPrinterList.Text = $"{T("PrinterList")} :";
            buttonAdd.Text = T("Add");
        }

        private void ButtonAdd_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBoxPrinterAdd.Text))
            {
                checkedListBoxPrinter.Items.Add(textBoxPrinterAdd.Text, true);
                SaveCheckedListBox();
                textBoxPrinterAdd.Clear();
                textBoxPrinterAdd.Focus();
            }
        }

        private void RemovePrinter(object sender, EventArgs e)
        {
            checkedListBoxPrinter.Items.Remove(checkedListBoxPrinter.SelectedItem);
            SaveCheckedListBox();
        }

        private void SaveCheckedListBox()
        {
            Variables.PrinterList = string.Join(";", checkedListBoxPrinter.Items.Cast<string>().Select(item => item.Trim()));
        }
    }
}
