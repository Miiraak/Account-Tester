using System.Data;

namespace AccountTester
{
    public partial class OptionDrives : Form
    {
        static string T(string key) => LangManager.Instance.Translate(key);

        public OptionDrives()
        {
            InitializeComponent();

            UpdateTexts();
            LangManager.Instance.LanguageChanged += UpdateTexts;

            string[] driveList = Variables.DrivesList.Split(';').Select(p => p.Trim()).Where(p => !string.IsNullOrEmpty(p)).ToArray();
            foreach (var drive in driveList)
            {
                checkedListBoxDrives.Items.Add(drive, true);
            }
        }

        private void UpdateTexts()
        {
            this.Text = T("DrivesOptions");
            labelDrivesList.Text = $"{T("DrivesList")} :";
            ButtonAdd.Text = T("Add");
        }

        private void ButtonAdd_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBoxDrivesAdd.Text))
            {
                checkedListBoxDrives.Items.Add(textBoxDrivesAdd.Text, true);
                SaveCheckedListBox();
                textBoxDrivesAdd.Clear();
                textBoxDrivesAdd.Focus();
            }
        }

        private void RemoveDrives(object sender, EventArgs e)
        {
            checkedListBoxDrives.Items.Remove(checkedListBoxDrives.SelectedItem);
            SaveCheckedListBox();
        }

        private void SaveCheckedListBox()
        {
            Variables.DrivesList = string.Join(";", checkedListBoxDrives.Items.Cast<string>().Select(item => item.Trim()));
        }
    }
}
