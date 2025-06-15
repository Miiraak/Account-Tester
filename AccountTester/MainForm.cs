using BlobPE;
using Microsoft.Win32;
using System.Diagnostics;
using System.Security.Principal;


namespace AccountTester
{
    public partial class MainForm : Form
    {
        static string T(string key) => LangManager.Instance.Translate(key);

        public MainForm()
        {
            InitializeComponent();
            richTextBoxLogs.Font = new Font("Consolas", 10);
            exportToolStripMenuItem.Enabled = false;

            Blob.RemoveUpdateFiles();

            UpdateTexts();
            LangManager.Instance.LanguageChanged += UpdateTexts;
        }

        internal void MainFormLoad(object sender, EventArgs e)
        {
            toolStripComboBoxExtensionByDefault.Text = Blob.Get("BaseExtension") ?? ".zip";

            string savedLanguage = Blob.Get("Langage") ?? "en-US";
            switch (savedLanguage)
            {
                case "en-US":
                    enUSToolStripMenuItem.Checked = true;
                    LangChangeCheck(enUSToolStripMenuItem);
                    LangManager.Instance.SetLanguage("en-US");
                    break;
                case "fr-FR":
                    frFRToolStripMenuItem.Checked = true;
                    LangChangeCheck(frFRToolStripMenuItem);
                    LangManager.Instance.SetLanguage("fr-FR");
                    break;
                default:
                    enUSToolStripMenuItem.Checked = true;
                    LangChangeCheck(enUSToolStripMenuItem);
                    LangManager.Instance.SetLanguage("en-US");
                    break;
            }

            TimeoutToolStripTextBox.Text = Blob.GetInt("Timeout").ToString();

            autoExportToolStripMenuItem.Checked = Blob.GetBool("AutoExport");
            autorunToolStripMenuItem.Checked = Blob.GetBool("Autorun");

            if (Variables.IsAutoRun)
            {
                Autorun();
                Application.Exit();
            }
        }

        private void UpdateTexts()
        {
            labelLogs.Text = T("Logs");
            copyToolStripMenuItem.Text = T("Copy");
            exportToolStripMenuItem.Text = T("Export");
            startToolStripMenuItem.Text = T("Start");
            optionsToolStripMenuItem.Text = T("Options");
            languageToolStripMenuItem.Text = T("Language");
            helpToolStripMenuItem.Text = T("Help");
            contactToolStripMenuItem.Text = T("Contact");
            reportsToolStripMenuItem.Text = T("Report");
            autoExportToolStripMenuItem.Text = T("AutoExport");
            extensionByDefaultToolStripMenuItem.Text = T("ExtensionByDefault");
            saveToolStripMenuItem.Text = T("Save");
            ResetToolStripMenuItem.Text = T("Reset");
            InternetToolStripMenuItem.Text = T("Internet");
            NetworkStorageToolStripMenuItem.Text = T("NetworkStorage");
            PrinterToolStripMenuItem.Text = T("Printer");
            TimeoutToolStripMenuItem.Text = $"{T("Timeout")} :";
        }

        /// <summary>
        /// Method for executing all tests sequentially, displaying the results in the richTextBoxLogs.
        /// </summary>
        /// <returns></returns>
        async Task ExecutionSequentielle()
        {
            Variables.General_TotalSuccess = 0;
            Variables.General_TotalTests = 0;
            exportToolStripMenuItem.Enabled = false;
            richTextBoxLogs.Clear();

            try
            {
                Stopwatch stopwatch = new();
                stopwatch.Start();
                richTextBoxLogs.AppendText($"AccountTester - {T("TestReport")}" + Environment.NewLine);
                richTextBoxLogs.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + Environment.NewLine);
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("Username")} :" + Environment.NewLine);
                richTextBoxLogs.AppendText($"- {Variables.General_UserName}" + Environment.NewLine);
                richTextBoxLogs.AppendText(Environment.NewLine);

                if (InternetToolStripMenuItem.Checked)
                {
                    richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                    richTextBoxLogs.AppendText($"#### {T("Internet")} :" + Environment.NewLine);
                    await Tests.InternetConnexionTest(richTextBoxLogs);
                    richTextBoxLogs.AppendText(Environment.NewLine);
                }

                if (NetworkStorageToolStripMenuItem.Checked)
                {
                    richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                    richTextBoxLogs.AppendText($"#### {T("NetworkStorageRights")} :" + Environment.NewLine);
                    Tests.NetworkStorageRightsTesting(richTextBoxLogs);
                    richTextBoxLogs.AppendText(Environment.NewLine);
                }

                if (OfficeToolStripMenuItem.Checked)
                {
                    richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                    richTextBoxLogs.AppendText($"#### {T("OfficeVersion")} :" + Environment.NewLine);
                    Tests.OfficeVersionTesting(richTextBoxLogs);
                    richTextBoxLogs.AppendText(Environment.NewLine);

                    if (Variables.WordIsInstalled)
                    {
                        richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                        richTextBoxLogs.AppendText($"#### {T("OfficeRights")} :" + Environment.NewLine);
                        Tests.OfficeWRTesting(richTextBoxLogs);
                        richTextBoxLogs.AppendText(Environment.NewLine);
                    }
                }

                if (PrinterToolStripMenuItem.Checked)
                {
                    richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                    richTextBoxLogs.AppendText($"#### {T("Printer")} :" + Environment.NewLine);
                    Tests.PrinterTesting(richTextBoxLogs);
                    richTextBoxLogs.AppendText(Environment.NewLine);
                }

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("TestsFinished")} :" + Environment.NewLine);
                stopwatch.Stop();
                richTextBoxLogs.AppendText($"- {T("TotalTimeElapsed")} : " + stopwatch.ElapsedMilliseconds + " ms" + Environment.NewLine);
                richTextBoxLogs.AppendText($"- {T("TotalSuccess")} : {Variables.General_TotalSuccess}/{Variables.General_TotalTests}");

                System.Media.SoundPlayer player = new(@"C:\Windows\Media\Windows Message Nudge.wav");
                player.Play();

                startToolStripMenuItem.Text = T("Restart");
                exportToolStripMenuItem.Enabled = true;

                Variables.General_Resume = richTextBoxLogs.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sequential Execution Error : " + Environment.NewLine + ex.Message);
            }
        }

        internal void ExportReport()
        {
            string fileName = $"{T("Report")}_{Environment.UserName}_{DateTime.Now:yyyyMMddHHmmss}";
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            switch (toolStripComboBoxExtensionByDefault.Text ?? ".zip")
            {
                case ".csv":
                    ExportForm.ExportToCSV(fileName, filePath);
                    break;
                case ".xml":
                    ExportForm.ExportToXML(fileName, filePath);
                    break;
                case ".json":
                    ExportForm.ExportToJSON(fileName, filePath);
                    break;
                case ".txt":
                    ExportForm.ExportToTxt(fileName, filePath);
                    break;
                case ".log":
                    ExportForm.ExportToLog(fileName, filePath);
                    break;
                case ".zip":
                    ExportForm.ExportToZip(fileName, filePath);
                    break;
            }
        }

        internal void Autorun()
        {
            Task.Run(() => ExecutionSequentielle()).Wait();
            ExportReport();
        }

        private void StartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            startToolStripMenuItem.Enabled = false;
            Task.Run(() => ExecutionSequentielle()).Wait();
            exportToolStripMenuItem.Enabled = true;
            startToolStripMenuItem.Enabled = true;

            if (autoExportToolStripMenuItem.Checked && toolStripComboBoxExtensionByDefault.Text != String.Empty)
            {
                ExportReport();
                MessageBox.Show($"{T("ExportForm_ButtonExport_MessageBox_Success")}.", $"{T("Export")}", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ExportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExportForm exportForm = new();
            exportForm.ShowDialog();
        }

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (richTextBoxLogs.Text.Length > 0)
                Clipboard.SetText(richTextBoxLogs.Text);
        }

        private void LangChangeCheck(object MenuStripItem)
        {
            foreach (ToolStripMenuItem item in languageToolStripMenuItem.DropDownItems)
            {
                if (item == MenuStripItem)
                    item.Checked = true;
                else
                    item.Checked = false;
            }
        }

        private void EnUSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LangManager.Instance.SetLanguage("en-US");
            LangChangeCheck(enUSToolStripMenuItem);
        }

        private void FrFRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LangManager.Instance.SetLanguage("fr-FR");
            LangChangeCheck(frFRToolStripMenuItem);
        }

        private void ContactToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ContactForm contactForm = new();
            contactForm.ShowDialog();
        }

        /// <summary>
        /// Handles the click event for the "Save" menu item, saving the current application settings.
        /// </summary>
        /// <remarks>This method saves the following settings: <list type="bullet"> <item><description>The
        /// selected language from the language menu.</description></item> <item><description>The default file extension
        /// from the extension dropdown.</description></item> <item><description>The auto-export
        /// preference.</description></item> <item><description>The autorun preference.</description></item> </list> The
        /// settings are persisted using the <c>Blob</c> storage mechanism.</remarks>
        /// <param name="sender">The source of the event, typically the "Save" menu item.</param>
        /// <param name="e">The event data associated with the click event.</param>
        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string selectedLanguage = languageToolStripMenuItem.DropDownItems.Cast<ToolStripMenuItem>().FirstOrDefault(item => item.Checked)?.Text ?? "en-US";
            string defaultExtension;
            if (toolStripComboBoxExtensionByDefault.SelectedItem != null)
                defaultExtension = toolStripComboBoxExtensionByDefault.Text;
            else
                defaultExtension = ".txt";
            bool autoExport = autoExportToolStripMenuItem.Checked;

            Blob.Set("Langage", selectedLanguage);
            Blob.Set("BaseExtension", defaultExtension);
            Blob.Set("Timeout", TimeoutToolStripTextBox.Text);
            Blob.Set("AutoExport", autoExport.ToString());
            Blob.Set("Autorun", autorunToolStripMenuItem.Checked.ToString());
            Blob.Save();
        }

        /// <summary>
        /// Toggles the application's autorun setting by adding or removing it from the system's startup registry.
        /// </summary>
        /// <remarks>This method requires the application to be running with administrative privileges. If
        /// the user does not have  the necessary permissions, a message box will prompt them to restart the application
        /// as an administrator.  When enabling autorun, the application is added to the Windows registry under 
        /// <c>HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run</c>. If the operation is successful, 
        /// the autorun setting is updated and saved. When disabling autorun, the corresponding registry entry is
        /// removed.  The method also updates the state of the associated menu item's checked property to reflect the
        /// current autorun status.</remarks>
        /// <param name="sender">The source of the event, typically the menu item that was clicked.</param>
        /// <param name="e">An <see cref="EventArgs"/> instance containing the event data.</param>
        private void AutorunToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool IsElevated = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
            if (!IsElevated)
            {
                // If the application is not running with administrative privileges, prompt the user to restart as admin.
                DialogResult result = MessageBox.Show(T("RunAsAdminRequired"), T("Attention"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    // Restart the application with administrative privileges.
                    ProcessStartInfo startInfo = new()
                    {
                        FileName = Application.ExecutablePath,
                        UseShellExecute = true,
                        Verb = "runas" // This will prompt for admin rights
                    };
                    Process.Start(startInfo);
                    Application.Exit();
                }
                return;
            }

            if (Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "AccountTester", null) == null)
            {
                if (MessageBox.Show(T("WantEnableAutorun"), T("Attention"), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    return;
                else
                {
                    Registry.SetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "AccountTester", $"\"{Application.ExecutablePath}\" --autorun");

                    if (Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "AccountTester", null) != null)
                    {
                        Blob.Set("Autorun", "true");
                        MessageBox.Show(T("AutorunEnabled"), T("Success"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                        autoExportToolStripMenuItem.Checked = false;
                        Blob.Set("AutoExport", "false");
                    }
                    else
                        MessageBox.Show(T("AutorunError"));
                }
            }
            else
            {
                using RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
                key.DeleteValue("AccountTester", false);
                Blob.Set("Autorun", "false");
                MessageBox.Show(T("AutorunDisabled"), T("Success"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            Blob.Save();
        }

        private void ClearFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Remove all related app files,
            // including update files, logs, and other temporary files.

            // Remove update files in temp
            // Remove logs 
            // Remove exported reports
            // Remove the registry entry for autorun if it exists
            // Reset the Blob storage
            // Remove other temporary files
        }

        private void TestsToolStripMenuItem_DropDownClosed(object sender, EventArgs e)
        {
            if (int.TryParse(TimeoutToolStripTextBox.Text, out int timeout) && timeout > 0)
                Variables.Timeout = timeout;
            else
                TimeoutToolStripTextBox.Text = Variables.Timeout.ToString();
        }
    }
}