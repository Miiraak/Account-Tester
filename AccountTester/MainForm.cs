using BlobPE;
using Microsoft.Win32;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Word = Microsoft.Office.Interop.Word;

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
        }

        /// <summary>
        /// Method for testing internet connection
        /// </summary>
        private async Task InternetConnexionTest()
        {
            Variables.General_TotalTests++;
            Stopwatch stopwatch = new();

            try
            {
                stopwatch.Start();
                using HttpClient client = new();
                client.Timeout = TimeSpan.FromSeconds(5);
                HttpResponseMessage response = await client.GetAsync(Variables.InternetConnexion_TestedURL);

                Variables.InternetConnexion_Hour = DateTime.Now.ToString("HH:mm:ss");
                Variables.InternetConnexion_HTMLStatut = response.StatusCode.ToString();

                if (response.IsSuccessStatusCode)
                {
                    richTextBoxLogs.AppendText($"{T("MainForm_RTBL_Internet_Connected")}" + Environment.NewLine);
                    Variables.General_TotalSuccess++;
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
            Variables.InternetConnexion_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>
        /// Method for testing network storage rights by checking if the user can read and write to network drives.
        /// </summary>
        private void NetworkStorageRightsTesting()
        {
            Variables.NetworkStorageRights_Hour = DateTime.Now.ToString("HH:mm:ss");
            Stopwatch stopwatch = new();
            stopwatch.Start();

            try
            {
                foreach (var drive in DriveInfo.GetDrives())
                {
                    Variables.General_TotalTests++;
                    Variables.NetworkStorageRights_DiskLetter = Variables.NetworkStorageRights_DiskLetter.Append(drive.Name).ToArray();

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
                                Variables.General_TotalSuccess++;
                                Variables.NetworkStorageRights_CheminUNC = Variables.NetworkStorageRights_CheminUNC.Append(cheminUNC).ToArray();
                                Variables.NetworkStorageRights_Serveur = Variables.NetworkStorageRights_Serveur.Append(serveur).ToArray();
                                Variables.NetworkStorageRights_ShareName = Variables.NetworkStorageRights_ShareName.Append(shareName).ToArray();
                            }

                            File.Delete(testFile);
                        }
                        catch (UnauthorizedAccessException)
                        {
                            richTextBoxLogs.AppendText($@"- {drive.Name} : {T("MainForm_RTBL_NetworkStorageRightsTesting_Refused")}" + Environment.NewLine);
                            Variables.NetworkStorageRights_CheminUNC = Variables.NetworkStorageRights_CheminUNC.Append(T("UnauthorizedAccess")).ToArray();
                            Variables.NetworkStorageRights_Serveur = Variables.NetworkStorageRights_Serveur.Append(T("UnauthorizedAccess")).ToArray();
                            Variables.NetworkStorageRights_ShareName = Variables.NetworkStorageRights_ShareName.Append(T("UnauthorizedAccess")).ToArray();
                        }
                        catch (IOException)
                        {
                            richTextBoxLogs.AppendText($@"- {drive.Name} : {T("MainForm_RTBL_NetworkStorageRightsTesting_Error")}" + Environment.NewLine);
                            Variables.NetworkStorageRights_CheminUNC = Variables.NetworkStorageRights_CheminUNC.Append(T("IOError")).ToArray();
                            Variables.NetworkStorageRights_Serveur = Variables.NetworkStorageRights_Serveur.Append(T("IOError")).ToArray();
                            Variables.NetworkStorageRights_ShareName = Variables.NetworkStorageRights_ShareName.Append(T("IOError")).ToArray();
                        }
                    }
                    else
                    {
                        Variables.NetworkStorageRights_CheminUNC = Variables.NetworkStorageRights_CheminUNC.Append(drive.Name).ToArray();
                        Variables.NetworkStorageRights_Serveur = Variables.NetworkStorageRights_Serveur.Append("localhost").ToArray();
                        Variables.NetworkStorageRights_ShareName = Variables.NetworkStorageRights_ShareName.Append(T("None")).ToArray();

                        richTextBoxLogs.AppendText($@"- {drive.Name} : {T("Omitted")}" + Environment.NewLine);
                        Variables.General_TotalSuccess++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error NetworkStorageRights : " + Environment.NewLine + ex);
            }

            stopwatch.Stop();
            Variables.NetworkStorageRights_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>
        /// Method for testing Office version on the system
        /// </summary>
        private void OfficeVersionTesting()
        {
            Variables.General_TotalTests++;
            Variables.OfficeVersion_Hour = DateTime.Now.ToString("HH:mm:ss");
            Stopwatch stopwatch = new();
            stopwatch.Start();

            try
            {
                using RegistryKey? key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0");
                string? officeVersion = key?.GetValue("OfficeProductReleaseIds")?.ToString();

                if (!string.IsNullOrEmpty(officeVersion))
                {
                    Variables.OfficeVersion_OfficeVersion = officeVersion;

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
                    Variables.General_TotalSuccess++;
                    Variables.WordIsInstalled = true;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("MainForm_RTBL_OfficeVersionTesting_NotFound")}" + Environment.NewLine);
                }

                Variables.OfficeVersion_OfficePath = GetRegValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "InstallationPath");
                Variables.OfficeVersion_OfficeCulture = GetRegValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0", "OfficeCulture");
                Variables.OfficeVersion_OfficeExcludedApps = GetRegValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0", "OfficeExcludedApps");
                Variables.OfficeVersion_OfficeLastUpdateStatus = GetRegValue(@"SOFTWARE\Microsoft\Office\ClickToRun\UpdateStatus", "LastUpdateResult");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error OfficeVersion : " + Environment.NewLine + ex.Message);
            }

            stopwatch.Stop();
            Variables.OfficeVersion_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
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
            Variables.General_TotalTests += 5;
            Variables.OfficeRights_Hour = DateTime.Now.ToString("HH:mm:ss");
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
                    Variables.OfficeRights_Create = "True";
                    Variables.General_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("Create")} : FAIL." + Environment.NewLine);
                    Variables.OfficeRights_Create = "False";
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
                    Variables.OfficeRights_Save = "True";
                    Variables.General_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("Save")} : FAIL" + Environment.NewLine);
                    Variables.OfficeRights_Save = "False";
                }
                doc.Close();

                doc = wordApp.Documents.Open(filePath);
                if (doc.Content.Text.Contains("The quick brown fox jumps over the lazy dog"))
                {
                    richTextBoxLogs.AppendText($"- {T("Read")} : OK" + Environment.NewLine);
                    richTextBoxLogs.AppendText($"- {T("Write")} : OK" + Environment.NewLine);
                    Variables.OfficeRights_Read = "True";
                    Variables.OfficeRights_Write = "True";
                    Variables.General_TotalSuccess += 2;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("Read")} : FAIL" + Environment.NewLine);
                    richTextBoxLogs.AppendText($"- {T("Write")} : FAIL" + Environment.NewLine);
                    Variables.OfficeRights_Read = "False";
                    Variables.OfficeRights_Write = "False";
                }
                doc.Close();

                wordApp.Quit();
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wordApp);

                File.Delete(filePath);
                if (!File.Exists(filePath))
                {
                    richTextBoxLogs.AppendText($"- {T("Delete")} : OK" + Environment.NewLine);
                    Variables.OfficeRights_Delete = "True";
                    Variables.General_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("Delete")} : FAIL" + Environment.NewLine);
                    Variables.OfficeRights_Delete = "False";
                }
                stopwatch.Stop();
                Variables.OfficeRights_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
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
            Variables.Printer_Hour = DateTime.Now.ToString("HH:mm:ss");
            Stopwatch stopwatch = new();
            stopwatch.Start();

            if (PrinterSettings.InstalledPrinters.Count == 0)
            {
                Variables.General_TotalTests++;
                richTextBoxLogs.AppendText(T("NoPrinterFound") + Environment.NewLine);
                Variables.General_TotalSuccess++;
                stopwatch.Stop();
                Variables.Printer_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
            }
            else
            {
                foreach (string printer in PrinterSettings.InstalledPrinters)
                {
                    if (!printer.Contains("Microsoft Print to PDF", StringComparison.OrdinalIgnoreCase) &&
                        !printer.Contains("XPS", StringComparison.OrdinalIgnoreCase) &&
                        !printer.Contains("OneNote", StringComparison.OrdinalIgnoreCase))
                    {
                        Variables.General_TotalTests++;
                        string registryPath = @"SYSTEM\CurrentControlSet\Control\Print\Printers\" + printer;

                        using RegistryKey? printerKey = Registry.LocalMachine.OpenSubKey(registryPath);
                        if (printerKey != null)
                        {
                            Variables.Printer_PrinterName = Variables.Printer_PrinterName.Append(printer).ToArray();
                            Variables.Printer_PrinterDriver = Variables.Printer_PrinterDriver.Append(printerKey.GetValue("Printer Driver").ToString()).ToArray();
                            Variables.Printer_PrinterPort = Variables.Printer_PrinterPort.Append(printerKey.GetValue("Port").ToString()).ToArray();

                            string? locationValue = printerKey.GetValue("Location")?.ToString();
                            if (!string.IsNullOrEmpty(locationValue))
                            {
                                string PrinterIP = locationValue.Split("//").Last().Split(":").First();
                                Variables.Printer_PrinterIP = Variables.Printer_PrinterIP.Append(PrinterIP).ToArray();

                                if (!string.IsNullOrEmpty(PrinterIP))
                                {
                                    Ping ping = new();
                                    PingReply reply = ping.Send(PrinterIP, 1000);

                                    if (reply.Status == IPStatus.Success)
                                    {
                                        richTextBoxLogs.AppendText(printer + Environment.NewLine);
                                        richTextBoxLogs.AppendText("- IP : " + PrinterIP + Environment.NewLine + "- Ping : OK" + Environment.NewLine);
                                        Variables.Printer_PrinterStatus = Variables.Printer_PrinterStatus.Append("OK").ToArray();
                                        Variables.General_TotalSuccess++;
                                    }
                                    else
                                    {
                                        richTextBoxLogs.AppendText(printer + Environment.NewLine);
                                        richTextBoxLogs.AppendText("- IP : " + PrinterIP + Environment.NewLine + "- Ping : FAIL" + Environment.NewLine);
                                        Variables.Printer_PrinterStatus = Variables.Printer_PrinterStatus.Append("FAIL").ToArray();
                                    }
                                }
                                else
                                {
                                    richTextBoxLogs.AppendText(printer + Environment.NewLine);
                                    richTextBoxLogs.AppendText($"- IP : {T("MainForm_RTBL_PrinterTesting_NotFound")}" + Environment.NewLine);
                                    Variables.Printer_PrinterIP = Variables.Printer_PrinterIP.Append(T("Unknown")).ToArray();
                                    Variables.Printer_PrinterStatus = Variables.Printer_PrinterStatus.Append(T("Unknown")).ToArray();
                                    Variables.Printer_PrinterDriver = Variables.Printer_PrinterDriver.Append(T("Unknown")).ToArray();
                                    Variables.Printer_PrinterPort = Variables.Printer_PrinterPort.Append(T("Unknown")).ToArray();
                                }
                            }
                            else
                            {
                                richTextBoxLogs.AppendText(printer + Environment.NewLine);
                                richTextBoxLogs.AppendText($"- {T("MainForm_RTBL_PrinterTesting_NoLocationValueReg")}" + Environment.NewLine);
                                Variables.Printer_PrinterIP = Variables.Printer_PrinterIP.Append(T("Unknown")).ToArray();
                                Variables.Printer_PrinterStatus = Variables.Printer_PrinterStatus.Append(T("Unknown")).ToArray();
                                Variables.Printer_PrinterDriver = Variables.Printer_PrinterDriver.Append(T("Unknown")).ToArray();
                                Variables.Printer_PrinterPort = Variables.Printer_PrinterPort.Append(T("Unknown")).ToArray();
                            }
                        }
                        else
                        {
                            richTextBoxLogs.AppendText(printer + Environment.NewLine);
                            richTextBoxLogs.AppendText($"- {T("MainForm_RTBL_NoRegKey")}" + Environment.NewLine);
                            Variables.Printer_PrinterIP = Variables.Printer_PrinterIP.Append(T("Unknown")).ToArray();
                            Variables.Printer_PrinterStatus = Variables.Printer_PrinterStatus.Append(T("Unknown")).ToArray();
                            Variables.Printer_PrinterDriver = Variables.Printer_PrinterDriver.Append(T("Unknown")).ToArray();
                            Variables.Printer_PrinterPort = Variables.Printer_PrinterPort.Append(T("Unknown")).ToArray();
                        }
                    }
                    else
                    {
                        richTextBoxLogs.AppendText($"{printer} : {T("Omitted")}" + Environment.NewLine);
                    }
                }
            }

            stopwatch.Stop();
            Variables.Printer_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
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

                if (Variables.WordIsInstalled)
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
                MessageBox.Show($"{T("ExportForm_ButtonExport_MessageBox_Success")}.");
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

        private void ChangeCheck(object MenuStripItem)
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
    }
}