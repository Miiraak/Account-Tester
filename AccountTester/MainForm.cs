using BlobPE;
using Microsoft.Win32;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace AccountTester
{
    public partial class MainForm : Form
    {
        static string T(string key) => LangManager.Instance.Translate(key);
        bool _WordIsInstalled = false;

        public MainForm()
        {
            InitializeComponent();
            Blob.RemoveUpdateFiles();
            richTextBoxLogs.Font = new Font("Consolas", 10);
            exportToolStripMenuItem.Enabled = false;

            UpdateTexts();
            LangManager.Instance.LanguageChanged += UpdateTexts;
        }

        public void MainFormLoad(object sender, EventArgs e)
        {
            toolStripComboBoxExtensionByDefault.Text = Blob.Get("BaseExtension");

            string savedLanguage = Blob.Get("Langage") ?? "en-US";
            switch (savedLanguage)
            {
                case "en-US":
                    enUSToolStripMenuItem.Checked = true;
                    ChangeCheck(enUSToolStripMenuItem);
                    LangManager.Instance.SetLanguage("en-US");
                    break;
                case "fr-FR":
                    frFRToolStripMenuItem.Checked = true;
                    ChangeCheck(frFRToolStripMenuItem);
                    LangManager.Instance.SetLanguage("fr-FR");
                    break;
                default:
                    enUSToolStripMenuItem.Checked = true;
                    ChangeCheck(enUSToolStripMenuItem);
                    LangManager.Instance.SetLanguage("en-US");
                    break;
            }

            autoExportToolStripMenuItem.Checked = Blob.GetBool("AutoExport");

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
        }

        /// <summary>
        /// Method for testing internet connection
        /// </summary>
        private async Task InternetConnexionTest()
        {
            ExportVariables.General_TotalTests++;
            Stopwatch stopwatch = new();

            try
            {
                stopwatch.Start();
                using HttpClient client = new();
                client.Timeout = TimeSpan.FromSeconds(5);
                HttpResponseMessage response = await client.GetAsync(ExportVariables.InternetConnexion_TestedURL);

                ExportVariables.InternetConnexion_Hour = DateTime.Now.ToString("HH:mm:ss");
                ExportVariables.InternetConnexion_HTMLStatut = response.StatusCode.ToString();

                if (response.IsSuccessStatusCode)
                {
                    richTextBoxLogs.AppendText($"{T("MainForm_RTBL_Internet_Connected")}" + Environment.NewLine);
                    ExportVariables.General_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText($"{T("MainForm_RTBL_Internet_Others")}" + response.StatusCode + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                richTextBoxLogs.AppendText($"{T("MainForm_RTBL_Internet_Others")}" + ex.InnerException?.Message + Environment.NewLine);
            }

            stopwatch.Stop();
            ExportVariables.InternetConnexion_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>
        /// Method for testing network storage rights by checking if the user can read and write to network drives.
        /// </summary>
        private void NetworkStorageRightsTesting()
        {
            ExportVariables.NetworkStorageRights_Hour = DateTime.Now.ToString("HH:mm:ss");
            Stopwatch stopwatch = new();
            stopwatch.Start();

            try
            {
                foreach (var drive in DriveInfo.GetDrives())
                {
                    ExportVariables.General_TotalTests++;
                    ExportVariables.NetworkStorageRights_DiskLetter = ExportVariables.NetworkStorageRights_DiskLetter.Append(drive.Name).ToArray();

                    if (drive.DriveType == DriveType.Network)
                    {
                        string cheminUNC = drive.RootDirectory.FullName;
                        string serveur = "";
                        string shareName = "";

                        var uncParts = cheminUNC.TrimEnd('\\').Split('\\');
                        if (uncParts.Length >= 4)
                        {
                            serveur = uncParts[2];
                            shareName = uncParts[3];
                        }
                        else
                        {
                            serveur = T("Unknown");
                            shareName = T("Unknown");
                        }

                        try
                        {
                            string testFile = Path.Combine(drive.RootDirectory.FullName, "test.txt");
                            File.WriteAllText(testFile, "test");

                            if (File.Exists(testFile))
                            {
                                richTextBoxLogs.AppendText($@"- {drive.Name} : OK" + Environment.NewLine);
                                ExportVariables.General_TotalSuccess++;
                                ExportVariables.NetworkStorageRights_CheminUNC = ExportVariables.NetworkStorageRights_CheminUNC.Append(cheminUNC).ToArray();
                                ExportVariables.NetworkStorageRights_Serveur = ExportVariables.NetworkStorageRights_Serveur.Append(serveur).ToArray();
                                ExportVariables.NetworkStorageRights_ShareName = ExportVariables.NetworkStorageRights_ShareName.Append(shareName).ToArray();
                            }

                            File.Delete(testFile);
                        }
                        catch (UnauthorizedAccessException)
                        {
                            richTextBoxLogs.AppendText($@"- {drive.Name} : {T("MainForm_RTBL_NetworkStorageRightsTesting_Refused")}" + Environment.NewLine);
                            ExportVariables.NetworkStorageRights_CheminUNC = ExportVariables.NetworkStorageRights_CheminUNC.Append(T("UnauthorizedAccess")).ToArray();
                            ExportVariables.NetworkStorageRights_Serveur = ExportVariables.NetworkStorageRights_Serveur.Append(T("UnauthorizedAccess")).ToArray();
                            ExportVariables.NetworkStorageRights_ShareName = ExportVariables.NetworkStorageRights_ShareName.Append(T("UnauthorizedAccess")).ToArray();
                        }
                        catch (IOException)
                        {
                            richTextBoxLogs.AppendText($@"- {drive.Name} : {T("MainForm_RTBL_NetworkStorageRightsTesting_Error")}" + Environment.NewLine);
                            ExportVariables.NetworkStorageRights_CheminUNC = ExportVariables.NetworkStorageRights_CheminUNC.Append(T("IOError")).ToArray();
                            ExportVariables.NetworkStorageRights_Serveur = ExportVariables.NetworkStorageRights_Serveur.Append(T("IOError")).ToArray();
                            ExportVariables.NetworkStorageRights_ShareName = ExportVariables.NetworkStorageRights_ShareName.Append(T("IOError")).ToArray();
                        }
                    }
                    else
                    {
                        ExportVariables.NetworkStorageRights_CheminUNC = ExportVariables.NetworkStorageRights_CheminUNC.Append(drive.Name).ToArray();
                        ExportVariables.NetworkStorageRights_Serveur = ExportVariables.NetworkStorageRights_Serveur.Append("localhost").ToArray();
                        ExportVariables.NetworkStorageRights_ShareName = ExportVariables.NetworkStorageRights_ShareName.Append(T("None")).ToArray();

                        richTextBoxLogs.AppendText($@"- {drive.Name} : {T("Omitted")}" + Environment.NewLine);
                        ExportVariables.General_TotalSuccess++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error NetworkStorageRights : " + Environment.NewLine + ex);
            }

            stopwatch.Stop();
            ExportVariables.NetworkStorageRights_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>
        /// Method for testing Office version on the system
        /// </summary>
        private void OfficeVersionTesting()
        {
            ExportVariables.General_TotalTests++;
            ExportVariables.OfficeVersion_Hour = DateTime.Now.ToString("HH:mm:ss");
            Stopwatch stopwatch = new();
            stopwatch.Start();

            try
            {
                using RegistryKey? key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0");
                string? officeVersion = key?.GetValue("OfficeProductReleaseIds")?.ToString();

                if (!string.IsNullOrEmpty(officeVersion))
                {
                    ExportVariables.OfficeVersion_OfficeVersion = officeVersion;

                    if (officeVersion.Contains(','))
                    {
                        foreach (string version in officeVersion.Split(','))
                        {
                            richTextBoxLogs.AppendText($"- {version}" + Environment.NewLine);
                        }
                    }
                    else
                    {
                        richTextBoxLogs.AppendText($"- {officeVersion}" + Environment.NewLine);
                    }
                    ExportVariables.General_TotalSuccess++;
                    _WordIsInstalled = true;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("MainForm_RTBL_OfficeVersionTesting_NotFound")}" + Environment.NewLine);
                }

                ExportVariables.OfficeVersion_OfficePath = GetRegValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "InstallationPath");
                ExportVariables.OfficeVersion_OfficeCulture = GetRegValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0", "OfficeCulture");
                ExportVariables.OfficeVersion_OfficeExcludedApps = GetRegValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0", "OfficeExcludedApps");
                ExportVariables.OfficeVersion_OfficeLastUpdateStatus = GetRegValue(@"SOFTWARE\Microsoft\Office\ClickToRun\UpdateStatus", "LastUpdateResult");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error OfficeVersion : " + Environment.NewLine + ex.Message);
            }

            stopwatch.Stop();
            ExportVariables.OfficeVersion_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        private static string GetRegValue(string path, string value)
        {
            using RegistryKey? regKey = Registry.LocalMachine.OpenSubKey(path);
            string? str = regKey?.GetValue(value)?.ToString();
            if (string.IsNullOrEmpty(str))
                return "Null";
            else
                return str;
        }

        /// <summary>                                      
        /// Method for testing Office Write and Read rights on the system by simulating a Word document creation and editing.
        /// </summary>
        private void OfficeWRTesting()
        {
            ExportVariables.General_TotalTests += 5;
            ExportVariables.OfficeRights_Hour = DateTime.Now.ToString("HH:mm:ss");
            Stopwatch stopwatch = new();

            try
            {
                stopwatch.Start();

                string fileName = $"temp_{Guid.NewGuid()}.doc";   // Guid named file to avoid collision.
                string filePath = Path.Combine(Path.GetTempPath(), fileName);
                Word.Application wordApp = new()
                {
                    Visible = false
                };

                Word.Document doc = wordApp.Documents.Add();
                doc.Content.Text = "The quick brown fox jumps over the lazy dog";
                doc.SaveAs2(filePath);
                doc.Close();
                if (File.Exists(filePath))
                {
                    richTextBoxLogs.AppendText($"- {T("Create")} : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_Create = "True";
                    ExportVariables.General_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("Create")} : FAIL." + Environment.NewLine);
                    ExportVariables.OfficeRights_Create = "False";
                    return;
                }

                doc = wordApp.Documents.Open(filePath);
                doc.Content.Text += "\nAdding more fox over the lazy dog.";
                doc.Save();
                doc.Close();

                doc = wordApp.Documents.Open(filePath);
                if (doc.Content.Text.Contains("Adding more fox over the lazy dog"))
                {
                    richTextBoxLogs.AppendText($"- {T("Save")} : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_Save = "True";
                    ExportVariables.General_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("Save")} : FAIL" + Environment.NewLine);
                    ExportVariables.OfficeRights_Save = "False";
                }
                doc.Close();

                doc = wordApp.Documents.Open(filePath);
                if (doc.Content.Text.Contains("The quick brown fox jumps over the lazy dog"))
                {
                    richTextBoxLogs.AppendText($"- {T("Read")} : OK" + Environment.NewLine);
                    richTextBoxLogs.AppendText($"- {T("Write")} : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_Read = "True";
                    ExportVariables.OfficeRights_Write = "True";
                    ExportVariables.General_TotalSuccess += 2;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("Read")} : FAIL" + Environment.NewLine);
                    richTextBoxLogs.AppendText($"- {T("Write")} : FAIL" + Environment.NewLine);
                    ExportVariables.OfficeRights_Read = "False";
                    ExportVariables.OfficeRights_Write = "False";
                }
                doc.Close();

                wordApp.Quit();
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wordApp);

                File.Delete(filePath);
                if (!File.Exists(filePath))
                {
                    richTextBoxLogs.AppendText($"- {T("Delete")} : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_Delete = "True";
                    ExportVariables.General_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("Delete")} : FAIL" + Environment.NewLine);
                    ExportVariables.OfficeRights_Delete = "False";
                }
                stopwatch.Stop();
                ExportVariables.OfficeRights_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error OfficeRights: " + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// Method for testing printers on the system
        /// </summary>
        private void PrinterTesting()
        {
            ExportVariables.Printer_Hour = DateTime.Now.ToString("HH:mm:ss");
            Stopwatch stopwatch = new();
            stopwatch.Start();

            if (PrinterSettings.InstalledPrinters.Count == 0)
            {
                ExportVariables.General_TotalTests++;
                richTextBoxLogs.AppendText(T("NoPrinterFound") + Environment.NewLine);
                ExportVariables.General_TotalSuccess++;
                stopwatch.Stop();
                ExportVariables.Printer_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
            }
            else
            {
                foreach (string printer in PrinterSettings.InstalledPrinters)
                {
                    if (!printer.Contains("Microsoft Print to PDF", StringComparison.OrdinalIgnoreCase) &&
                        !printer.Contains("XPS", StringComparison.OrdinalIgnoreCase) &&
                        !printer.Contains("OneNote", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportVariables.General_TotalTests++;
                        string registryPath = @"SYSTEM\CurrentControlSet\Control\Print\Printers\" + printer;

                        using RegistryKey? printerKey = Registry.LocalMachine.OpenSubKey(registryPath);
                        if (printerKey != null)
                        {
                            ExportVariables.Printer_PrinterName = ExportVariables.Printer_PrinterName.Append(printer).ToArray();
                            ExportVariables.Printer_PrinterDriver = ExportVariables.Printer_PrinterDriver.Append(printerKey.GetValue("Printer Driver").ToString()).ToArray();
                            ExportVariables.Printer_PrinterPort = ExportVariables.Printer_PrinterPort.Append(printerKey.GetValue("Port").ToString()).ToArray();

                            string? locationValue = printerKey.GetValue("Location")?.ToString();
                            if (!string.IsNullOrEmpty(locationValue))
                            {
                                string PrinterIP = locationValue.Split("//").Last().Split(":").First();
                                ExportVariables.Printer_PrinterIP = ExportVariables.Printer_PrinterIP.Append(PrinterIP).ToArray();

                                if (!string.IsNullOrEmpty(PrinterIP))
                                {
                                    Ping ping = new();
                                    PingReply reply = ping.Send(PrinterIP, 1000);

                                    if (reply.Status == IPStatus.Success)
                                    {
                                        richTextBoxLogs.AppendText(printer + Environment.NewLine);
                                        richTextBoxLogs.AppendText("- IP : " + PrinterIP + Environment.NewLine + "- Ping : OK" + Environment.NewLine);
                                        ExportVariables.Printer_PrinterStatus = ExportVariables.Printer_PrinterStatus.Append("OK").ToArray();
                                        ExportVariables.General_TotalSuccess++;
                                    }
                                    else
                                    {
                                        richTextBoxLogs.AppendText(printer + Environment.NewLine);
                                        richTextBoxLogs.AppendText("- IP : " + PrinterIP + Environment.NewLine + "- Ping : FAIL" + Environment.NewLine);
                                        ExportVariables.Printer_PrinterStatus = ExportVariables.Printer_PrinterStatus.Append("FAIL").ToArray();
                                    }
                                }
                                else
                                {
                                    richTextBoxLogs.AppendText(printer + Environment.NewLine);
                                    richTextBoxLogs.AppendText($"- IP : {T("MainForm_RTBL_PrinterTesting_NotFound")}" + Environment.NewLine);
                                    ExportVariables.Printer_PrinterIP = ExportVariables.Printer_PrinterIP.Append(T("Unknown")).ToArray();
                                    ExportVariables.Printer_PrinterStatus = ExportVariables.Printer_PrinterStatus.Append(T("Unknown")).ToArray();
                                    ExportVariables.Printer_PrinterDriver = ExportVariables.Printer_PrinterDriver.Append(T("Unknown")).ToArray();
                                    ExportVariables.Printer_PrinterPort = ExportVariables.Printer_PrinterPort.Append(T("Unknown")).ToArray();
                                }
                            }
                            else
                            {
                                richTextBoxLogs.AppendText(printer + Environment.NewLine);
                                richTextBoxLogs.AppendText($"- {T("MainForm_RTBL_PrinterTesting_NoLocationValueReg")}" + Environment.NewLine);
                                ExportVariables.Printer_PrinterIP = ExportVariables.Printer_PrinterIP.Append(T("Unknown")).ToArray();
                                ExportVariables.Printer_PrinterStatus = ExportVariables.Printer_PrinterStatus.Append(T("Unknown")).ToArray();
                                ExportVariables.Printer_PrinterDriver = ExportVariables.Printer_PrinterDriver.Append(T("Unknown")).ToArray();
                                ExportVariables.Printer_PrinterPort = ExportVariables.Printer_PrinterPort.Append(T("Unknown")).ToArray();
                            }
                        }
                        else
                        {
                            richTextBoxLogs.AppendText(printer + Environment.NewLine);
                            richTextBoxLogs.AppendText($"- {T("MainForm_RTBL_NoRegKey")}" + Environment.NewLine);
                            ExportVariables.Printer_PrinterIP = ExportVariables.Printer_PrinterIP.Append(T("Unknown")).ToArray();
                            ExportVariables.Printer_PrinterStatus = ExportVariables.Printer_PrinterStatus.Append(T("Unknown")).ToArray();
                            ExportVariables.Printer_PrinterDriver = ExportVariables.Printer_PrinterDriver.Append(T("Unknown")).ToArray();
                            ExportVariables.Printer_PrinterPort = ExportVariables.Printer_PrinterPort.Append(T("Unknown")).ToArray();
                        }
                    }
                    else
                    {
                        richTextBoxLogs.AppendText($"{printer} : {T("Omitted")}" + Environment.NewLine);
                    }
                }
            }

            stopwatch.Stop();
            ExportVariables.Printer_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>
        /// For output formatting, see the ExecutionSequentielle method instead.
        /// </summary>
        private void ButtonStart_Click(object sender, EventArgs e)
        {
            Task.Run(() => ExecutionSequentielle()).Wait();
            exportToolStripMenuItem.Enabled = true;
        }

        /// <summary>
        /// Method for executing all tests sequentially, displaying the results in the richTextBoxLogs.
        /// </summary>
        /// <returns></returns>
        async Task ExecutionSequentielle()
        {
            ExportVariables.General_TotalSuccess = 0;
            ExportVariables.General_TotalTests = 0;
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
                richTextBoxLogs.AppendText($"- {ExportVariables.General_UserName}" + Environment.NewLine);
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("Internet")} :" + Environment.NewLine);
                await InternetConnexionTest();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("NetworkStorageRights")} :" + Environment.NewLine);
                NetworkStorageRightsTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("OfficeVersion")} :" + Environment.NewLine);
                OfficeVersionTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                if (_WordIsInstalled)
                {
                    richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                    richTextBoxLogs.AppendText($"#### {T("OfficeRights")} :" + Environment.NewLine);
                    OfficeWRTesting();
                    richTextBoxLogs.AppendText(Environment.NewLine);
                }

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("Printer")} :" + Environment.NewLine);
                PrinterTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("TestsFinished")} :" + Environment.NewLine);
                stopwatch.Stop();
                richTextBoxLogs.AppendText($"- {T("TotalTimeElapsed")} : " + stopwatch.ElapsedMilliseconds + " ms" + Environment.NewLine);
                richTextBoxLogs.AppendText($"- {T("TotalSuccess")} : {ExportVariables.General_TotalSuccess}/{ExportVariables.General_TotalTests}");

                System.Media.SoundPlayer player = new(@"C:\Windows\Media\Windows Message Nudge.wav");
                player.Play();

                startToolStripMenuItem.Text = T("Restart");
                exportToolStripMenuItem.Enabled = true;

                ExportVariables.General_Resume = richTextBoxLogs.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sequential Execution Error : " + Environment.NewLine + ex.Message);
            }
        }

        private void EnUSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LangManager.Instance.SetLanguage("en-US");
            ChangeCheck(enUSToolStripMenuItem);
        }

        private void FrFRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LangManager.Instance.SetLanguage("fr-FR");
            ChangeCheck(frFRToolStripMenuItem);
        }

        private void ContactToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ContactForm contactForm = new();
            contactForm.ShowDialog();
        }

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (richTextBoxLogs.Text.Length > 0)
                Clipboard.SetText(richTextBoxLogs.Text);
        }

        private void ChangeCheck(object MenuStipItem)
        {
            foreach (ToolStripMenuItem item in languageToolStripMenuItem.DropDownItems)
            {
                if (item == MenuStipItem)
                    item.Checked = true;
                else
                    item.Checked = false;
            }
        }

        private void ExportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ExportForm exportForm = new();
            exportForm.ShowDialog();
        }

        private void StartToolStripMenuItem_Click(object sender, EventArgs e)
        {
            startToolStripMenuItem.Enabled = false;
            Task.Run(() => ExecutionSequentielle()).Wait();
            exportToolStripMenuItem.Enabled = true;
            startToolStripMenuItem.Enabled = true;

            if (autoExportToolStripMenuItem.Checked && toolStripComboBoxExtensionByDefault.Text != String.Empty)
                ExportReport();
        }

        internal void ExportReport()
        {
            string fileName = $"{T("Report")}_{Environment.UserName}_{DateTime.Now:yyyyMMddHHmmss}";
            string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string extension;

            switch (toolStripComboBoxExtensionByDefault.Text)
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

            MessageBox.Show($"{T("ExportForm_ButtonExport_MessageBox_Success")}.");
        }

        private void SaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Retrieve the texte of the currently selected toolstrip menu item from languageToolStripMenuItem
            string selectedLanguage = languageToolStripMenuItem.DropDownItems.Cast<ToolStripMenuItem>().FirstOrDefault(item => item.Checked)?.Text ?? "en-US";
            string defaultExtension;
            if (toolStripComboBoxExtensionByDefault.SelectedItem != null)
                defaultExtension = toolStripComboBoxExtensionByDefault.Text;
            else
                defaultExtension = ".txt";
            bool autoExport = autoExportToolStripMenuItem.Checked;

            Blob.Set("Langage", selectedLanguage);
            Blob.Set("BaseExtension", defaultExtension);
            Blob.Set("AutoExport", autoExport.ToString());
            Blob.Save();
        }
    }
}