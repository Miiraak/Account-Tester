using System.Diagnostics;
using System.Drawing.Printing;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Word = Microsoft.Office.Interop.Word;

namespace AccountTester
{
    public partial class MainForm : Form
    {
        string T(string key) => LangManager.Instance.Translate(key);
        bool _WordIsInstalled = false;

        public MainForm()
        {
            InitializeComponent();
            richTextBoxLogs.Font = new Font("Consolas", 10);
            buttonExportForm.Enabled = false;

            UpdateTexts();
            LangManager.Instance.LanguageChanged += UpdateTexts;
        }

        private void UpdateTexts()
        {
            labelLogs.Text = T("LabelLogs");
            buttonCopier.Text = T("CopyButton");
            buttonExportForm.Text = T("ExportButton");
            buttonStart.Text = T("StartButton");
            optionsToolStripMenuItem.Text = T("optionsToolStripMenuItem");
            langageToolStripMenuItem.Text = T("langageToolStripMenuItem");
        }

        /// <summary>
        /// Method for testing internet connection
        /// </summary>
        private async Task InternetConnexionTest()
        {
            ExportVariables.General_export_TotalTests++;
            Stopwatch stopwatch = new();

            try
            {
                stopwatch.Start();
                using HttpClient client = new();
                client.Timeout = TimeSpan.FromSeconds(5);
                HttpResponseMessage response = await client.GetAsync(ExportVariables.InternetConnexion_export_TestedURL);

                ExportVariables.InternetConnexion_export_Hour = DateTime.Now.ToString("HH:mm:ss");
                ExportVariables.InternetConnexion_export_HTMLStatut = response.StatusCode.ToString();

                if (response.IsSuccessStatusCode)
                {
                    richTextBoxLogs.AppendText($"{T("RTBL_Internet_1")}" + Environment.NewLine);
                    ExportVariables.General_export_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText($"{"RTBL_Internet_2"}" + response.StatusCode + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                richTextBoxLogs.AppendText($"{"RTBL_Internet_2"}" + ex.InnerException?.Message + Environment.NewLine);
            }

            stopwatch.Stop();
            ExportVariables.InternetConnexion_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>
        /// Method for testing network storage rights on the system by trying to write a file on each network drive.
        /// </summary>
        private void NetworkStorageRightsTesting()
        {
            ExportVariables.NetworkStorageRights_export_Hour = DateTime.Now.ToString("HH:mm:ss");
            Stopwatch stopwatch = new();
            stopwatch.Start();

            try
            {
                foreach (var drive in DriveInfo.GetDrives())
                {
                    ExportVariables.General_export_TotalTests++;

                    if (drive.DriveType == DriveType.Network)
                    {
                        string UNCpath = T("nullValue");
                        using (RegistryKey? key = Registry.CurrentUser.OpenSubKey("Network\\" + drive.Name[0].ToString()))
                            if (key != null)
                            {
                                UNCpath = key.GetValue("RemotePath")?.ToString() + drive.Name[2..].ToString();
                            }

                        ExportVariables.NetworkStorageRights_export_CheminUNC ??= [];
                        ExportVariables.NetworkStorageRights_export_CheminUNC = ExportVariables.NetworkStorageRights_export_CheminUNC.Append(UNCpath).ToArray();

                        try
                        {
                            string testFile = Path.Combine(drive.RootDirectory.FullName, "test.txt");
                            File.WriteAllText(testFile, "test");

                            if (File.Exists(testFile))
                            {
                                richTextBoxLogs.AppendText($@"- {drive.Name} : OK" + Environment.NewLine);
                                ExportVariables.General_export_TotalSuccess++;
                            }

                            File.Delete(testFile);
                        }
                        catch (UnauthorizedAccessException)
                        {
                            richTextBoxLogs.AppendText($@"- {drive.Name} : {T("RTBL_NetworkStorageRightsTesting_1")}" + Environment.NewLine);
                        }
                        catch (IOException)
                        {
                            richTextBoxLogs.AppendText($@"- {drive.Name} : {T("RTBL_NetworkStorageRightsTesting_2")}" + Environment.NewLine);
                        }
                    }
                    else
                    {
                        ExportVariables.NetworkStorageRights_export_DiskLetter ??= [];
                        ExportVariables.NetworkStorageRights_export_DiskLetter = ExportVariables.NetworkStorageRights_export_DiskLetter.Append(drive.Name).ToArray();

                        richTextBoxLogs.AppendText($@"- {drive.Name} : {T("RTBL_NetworkStorageRightsTesting_3")}" + Environment.NewLine);
                        ExportVariables.General_export_TotalSuccess++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error NetworkStorageRights : " + Environment.NewLine + ex);
            }

            stopwatch.Stop();
            ExportVariables.NetworkStorageRights_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>
        /// Method for testing Office version on the system
        /// </summary>
        private void OfficeVersionTesting()
        {
            ExportVariables.General_export_TotalTests++;
            ExportVariables.OfficeVersion_export_Hour = DateTime.Now.ToString("HH:mm:ss");
            Stopwatch stopwatch = new();
            stopwatch.Start();

            try
            {
                string registryPath = @"SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0";
                using RegistryKey? key = Registry.LocalMachine.OpenSubKey(registryPath);
                string? officeVersion = key?.GetValue("OfficeProductReleaseIds")?.ToString();

                if (!string.IsNullOrEmpty(officeVersion))
                {
                    ExportVariables.OfficeVersion_export_OfficeVersion = officeVersion;

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
                    ExportVariables.General_export_TotalSuccess++;
                    _WordIsInstalled = true;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("RTBL_OfficeVersionTesting_1")}" + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error OfficeVersion : " + Environment.NewLine + ex.Message);
            }

            stopwatch.Stop();
            ExportVariables.OfficeVersion_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>                                      
        /// Method for testing Office Write and Read rights on the system by simulating a Word document creation and editing.
        /// </summary>
        private void OfficeWRTesting()
        {
            ExportVariables.General_export_TotalTests += 5;
            ExportVariables.OfficeRights_export_Hour = DateTime.Now.ToString("HH:mm:ss");
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
                    richTextBoxLogs.AppendText($"- {T("RTBL_OfficeWRTesting_1")} : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanCreate = "True";
                    ExportVariables.General_export_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("RTBL_OfficeWRTesting_1")} : FAIL." + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanCreate = "False";
                    return;
                }

                doc = wordApp.Documents.Open(filePath);
                doc.Content.Text += "\nAdding more fox over the lazy dog.";
                doc.Save();
                doc.Close();

                doc = wordApp.Documents.Open(filePath);
                if (doc.Content.Text.Contains("Adding more fox over the lazy dog"))
                {
                    richTextBoxLogs.AppendText($"- {T("RTBL_OfficeWRTesting_2")} : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanSave = "True";
                    ExportVariables.General_export_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("RTBL_OfficeWRTesting_2")} : FAIL" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanSave = "False";
                }
                doc.Close();

                doc = wordApp.Documents.Open(filePath);
                if (doc.Content.Text.Contains("The quick brown fox jumps over the lazy dog"))
                {
                    richTextBoxLogs.AppendText($"- {T("RTBL_OfficeWRTesting_3")} : OK" + Environment.NewLine);
                    richTextBoxLogs.AppendText($"- {T("RTBL_OfficeWRTesting_4")} : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanRead = "True";
                    ExportVariables.OfficeRights_export_CanWrite = "True";
                    ExportVariables.General_export_TotalSuccess += 2;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("RTBL_OfficeWRTesting_3")} : FAIL" + Environment.NewLine);
                    richTextBoxLogs.AppendText($"- {T("RTBL_OfficeWRTesting_4")} : FAIL" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanRead = "False";
                    ExportVariables.OfficeRights_export_CanWrite = "False";
                }
                doc.Close();

                wordApp.Quit();
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wordApp);

                File.Delete(filePath);
                if (!File.Exists(filePath))
                {
                    richTextBoxLogs.AppendText($"- {T("RTBL_OfficeWRTesting_5")} : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanDelete = "True";
                    ExportVariables.General_export_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText($"- {T("RTBL_OfficeWRTesting_5")} : FAIL" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanDelete = "False";
                }
                stopwatch.Stop();
                ExportVariables.OfficeRights_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
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
            ExportVariables.Printer_export_Hour = DateTime.Now.ToString("HH:mm:ss");
            Stopwatch stopwatch = new();
            stopwatch.Start();

            if (PrinterSettings.InstalledPrinters.Count == 0)
            {
                ExportVariables.General_export_TotalTests++;
                richTextBoxLogs.AppendText($"{T("RTBL_PrinterTesting_1")}" + Environment.NewLine);
                ExportVariables.General_export_TotalSuccess++;
                stopwatch.Stop();
                ExportVariables.Printer_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
            }
            else
            {
                foreach (string printer in PrinterSettings.InstalledPrinters)
                {
                    if (!printer.Contains("Microsoft Print to PDF", StringComparison.OrdinalIgnoreCase) &&
                        !printer.Contains("XPS", StringComparison.OrdinalIgnoreCase) &&
                        !printer.Contains("OneNote", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportVariables.General_export_TotalTests++;
                        string registryPath = @"SYSTEM\CurrentControlSet\Control\Print\Printers\" + printer;

                        using RegistryKey? printerKey = Registry.LocalMachine.OpenSubKey(registryPath);
                        if (printerKey != null)
                        {
                            string? locationValue = printerKey.GetValue("Location")?.ToString();
                            if (!string.IsNullOrEmpty(locationValue))
                            {
                                string PrinterIP = locationValue.Split("//").Last().Split(":").First();

                                if (!string.IsNullOrEmpty(PrinterIP))
                                {
                                    Ping ping = new();
                                    PingReply reply = ping.Send(PrinterIP, 1000);

                                    if (reply.Status == IPStatus.Success)
                                    {
                                        richTextBoxLogs.AppendText(printer + Environment.NewLine);
                                        richTextBoxLogs.AppendText("- IP : " + PrinterIP + Environment.NewLine + "- Ping : OK" + Environment.NewLine);
                                        ExportVariables.General_export_TotalSuccess++;
                                    }
                                    else
                                    {
                                        richTextBoxLogs.AppendText(printer + Environment.NewLine);
                                        richTextBoxLogs.AppendText("- IP : " + PrinterIP + Environment.NewLine + "- Ping : FAIL" + Environment.NewLine);
                                    }
                                }
                                else
                                {
                                    richTextBoxLogs.AppendText(printer + Environment.NewLine);
                                    richTextBoxLogs.AppendText($"- IP : {T("RTBL_PrinterTesting_2")}" + Environment.NewLine);
                                }
                            }
                            else
                            {
                                richTextBoxLogs.AppendText(printer + Environment.NewLine);
                                richTextBoxLogs.AppendText($"- {T("RTBL_PrinterTesting_3")}" + Environment.NewLine);
                            }
                        }
                        else
                        {
                            richTextBoxLogs.AppendText(printer + Environment.NewLine);
                            richTextBoxLogs.AppendText($"- {T("RTBL_PrinterTesting_4")}" + Environment.NewLine);
                        }
                    }
                }
            }

            stopwatch.Stop();
            ExportVariables.Printer_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>
        /// For output formatting, see the ExecutionSequentielle method instead.
        /// </summary>
        private void ButtonStart_Click(object sender, EventArgs e)
        {
            buttonStart.Enabled = false;
            Task.Run(() => ExecutionSequentielle()).Wait();
            buttonCopier.Enabled = true;
            buttonExportForm.Enabled = true;
            buttonStart.Enabled = true;
        }

        /// <summary>
        /// Method for executing all tests sequentially, displaying the results in the richTextBoxLogs.
        /// </summary>
        /// <returns></returns>
        async Task ExecutionSequentielle()
        {
            ExportVariables.General_export_TotalSuccess = 0;
            ExportVariables.General_export_TotalTests = 0;
            buttonCopier.Enabled = false;
            buttonExportForm.Enabled = false;
            richTextBoxLogs.Clear();

            try
            {
                Stopwatch stopwatch = new();
                stopwatch.Start();
                richTextBoxLogs.AppendText($"AccountTester - {T("RTBL_ExecutionSequentielle_1")}" + Environment.NewLine);
                richTextBoxLogs.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + Environment.NewLine);
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("RTBL_ExecutionSequentielle_2")} :" + Environment.NewLine);
                richTextBoxLogs.AppendText($"- {ExportVariables.General_export_UserName}" + Environment.NewLine);
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("RTBL_ExecutionSequentielle_3")} :" + Environment.NewLine);
                await InternetConnexionTest();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("RTBL_ExecutionSequentielle_4")} :" + Environment.NewLine);
                NetworkStorageRightsTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("RTBL_ExecutionSequentielle_5")} :" + Environment.NewLine);
                OfficeVersionTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                if (_WordIsInstalled)
                {
                    richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                    richTextBoxLogs.AppendText($"#### {T("RTBL_ExecutionSequentielle_6")} :" + Environment.NewLine);
                    OfficeWRTesting();
                    richTextBoxLogs.AppendText(Environment.NewLine);
                }

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("RTBL_ExecutionSequentielle_7")} :" + Environment.NewLine);
                PrinterTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### {T("RTBL_ExecutionSequentielle_8")} :" + Environment.NewLine);
                stopwatch.Stop();
                richTextBoxLogs.AppendText($"- {T("RTBL_ExecutionSequentielle_9")} : " + stopwatch.ElapsedMilliseconds + " ms" + Environment.NewLine);
                richTextBoxLogs.AppendText($"- {T("RTBL_ExecutionSequentielle_10")} : {ExportVariables.General_export_TotalSuccess}/{ExportVariables.General_export_TotalTests}");
                ExportVariables.General_export_TotalSuccess = 0;
                ExportVariables.General_export_TotalTests = 0;

                System.Media.SoundPlayer player = new(@"C:\Windows\Media\Windows Message Nudge.wav");
                player.Play();

                buttonStart.Text = T("RTBL_ExecutionSequentielle_11");

                ExportVariables.General_export_Resume = richTextBoxLogs.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sequential Execution Error : " + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// Event handler for the Export button click event.
        /// </summary>
        private void ButtonExport_Click(object sender, EventArgs e)
        {
            ExportForm exportForm = new();
            exportForm.ShowDialog();
        }

        private void ButtonCopier_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBoxLogs.Text);
        }

        private void enUSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LangManager.Instance.SetLanguage("en-US");
        }

        private void frFRToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LangManager.Instance.SetLanguage("fr-FR");
        }
    }
}
