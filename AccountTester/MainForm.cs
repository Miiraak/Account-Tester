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
        bool _WordIsInstalled = false;

        public MainForm()
        {
            InitializeComponent();
            richTextBoxLogs.Font = new Font("Consolas", 10);
            buttonExportForm.Enabled = false;
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

                ExportVariables.InternetConnexion_export_DateAndHour = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                ExportVariables.InternetConnexion_export_HTMLStatut = response.StatusCode.ToString();

                if (response.IsSuccessStatusCode)
                {
                    richTextBoxLogs.AppendText("- État : Connecté" + Environment.NewLine);
                    ExportVariables.General_export_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText("- État : " + response.StatusCode + Environment.NewLine);
                }
                stopwatch.Stop();
                ExportVariables.InternetConnexion_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error InternetConnexion : " + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// Method for testing network storage rights on the system by trying to write a file on each network drive.
        /// </summary>
        private void NetworkStorageRightsTesting()
        {
            ExportVariables.NetworkStorageRights_export_DateAndHour = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Stopwatch stopwatch = new();
            stopwatch.Start();

            try
            {
                foreach (var drive in DriveInfo.GetDrives())
                {
                    ExportVariables.General_export_TotalTests++;

                    if (drive.DriveType == DriveType.Network)
                    {
                        try
                        {
                            string testFile = Path.Combine(drive.RootDirectory.FullName, "test.txt");
                            File.WriteAllText(testFile, "test");
                            File.Delete(testFile);
                            richTextBoxLogs.AppendText($@"- {drive.Name}\ : OK" + Environment.NewLine);
                            ExportVariables.General_export_TotalSuccess++;
                        }
                        catch (UnauthorizedAccessException)
                        {
                            richTextBoxLogs.AppendText($@"- {drive.Name}\ : Ecriture refusée" + Environment.NewLine);
                        }
                        catch (IOException)
                        {
                            richTextBoxLogs.AppendText($@"- {drive.Name}\ : Erreur connexion" + Environment.NewLine);
                        }
                    }
                    else
                    {
                        richTextBoxLogs.AppendText($@"- {drive.Name}\ : Omis" + Environment.NewLine);
                        ExportVariables.General_export_TotalSuccess++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur NetworkStorageRights : " + Environment.NewLine + ex.Message);
            }

            stopwatch.Stop();
            ExportVariables.InternetConnexion_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>
        /// Method for testing Office version on the system
        /// </summary>
        private void OfficeVersionTesting()
        {
            ExportVariables.General_export_TotalTests++;
            ExportVariables.OfficeVersion_export_DateAndHour = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
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
                    return;
                }

                stopwatch.Stop();
                ExportVariables.OfficeVersion_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
                richTextBoxLogs.AppendText("- Aucune version trouvée" + Environment.NewLine);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur OfficeVersion : " + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>                                      
        /// Method for testing Office Write and Read rights on the system by simulating a Word document creation and editing.
        /// </summary>
        private void OfficeWRTesting()
        {
            ExportVariables.General_export_TotalTests += 5;
            ExportVariables.OfficeRights_export_DateAndHour = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            ExportVariables.OfficeRights_export_FolderTested = [Path.GetTempPath()];
            Stopwatch stopwatch = new();

            try
            {
                stopwatch.Start();

                string fileName = Guid.NewGuid().ToString() + ".doc";   // Guid named file to avoid collision.
                string filePath = Path.Combine(Path.GetTempPath(), fileName);
                Word.Application wordApp = new()
                {
                    Visible = false
                };

                // Création du document
                Word.Document doc = wordApp.Documents.Add();
                doc.Content.Text = "Ceci est un test d'écriture dans un fichier .docx via Word.";
                doc.SaveAs2(filePath);
                doc.Close();
                if (File.Exists(filePath))
                {
                    richTextBoxLogs.AppendText("- Création : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanCreate = "True";
                    ExportVariables.General_export_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText("- Création : ECHEC." + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanCreate = "False";
                    return;
                }

                // Sauvegarde
                doc = wordApp.Documents.Open(filePath);
                doc.Content.Text += "\nAjout de texte pour le test de sauvegarde.";
                doc.Save();
                doc.Close();


                // Check if the content was added
                doc = wordApp.Documents.Open(filePath);
                if (doc.Content.Text.Contains("Ajout de texte pour le test de sauvegarde"))
                {
                    richTextBoxLogs.AppendText("- Sauvegarde : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanSave = "True";
                    ExportVariables.General_export_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText("- Sauvegarde : ECHEC" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanSave = "False";
                }
                doc.Close();

                // Réouverture du document
                doc = wordApp.Documents.Open(filePath);
                if (doc.Content.Text.Contains("Ceci est un test d'écriture dans un fichier .docx via Word."))
                {
                    richTextBoxLogs.AppendText("- Lecture : OK" + Environment.NewLine);
                    richTextBoxLogs.AppendText("- Ecriture : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanRead = "True";
                    ExportVariables.OfficeRights_export_CanWrite = "True";
                    ExportVariables.General_export_TotalSuccess += 2;

                }
                else
                {
                    richTextBoxLogs.AppendText("- Lecture : Echec" + Environment.NewLine);
                    richTextBoxLogs.AppendText("- Ecriture : Echec" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanRead = "False";
                    ExportVariables.OfficeRights_export_CanWrite = "False";

                }
                doc.Close();

                // Fermeture et nettoyage
                wordApp.Quit();
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wordApp);

                // Suppression du fichier
                File.Delete(filePath);
                if (!File.Exists(filePath))
                {
                    richTextBoxLogs.AppendText("- Suppression : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanDelete = "True";
                    ExportVariables.General_export_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText("- Suppression : Echec" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanDelete = "False";
                }

                stopwatch.Stop();
                ExportVariables.OfficeRights_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur OfficeRigts: " + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// Method for testing printers on the system
        /// </summary>
        private void PrinterTesting()
        {
            Stopwatch stopwatch = new();
            stopwatch.Start();

            // Gather all printers on the system
            if (PrinterSettings.InstalledPrinters.Count == 0)
            {
                ExportVariables.General_export_TotalTests++;
                richTextBoxLogs.AppendText("Aucune imprimante trouvée." + Environment.NewLine);
                ExportVariables.General_export_TotalSuccess++;
                stopwatch.Stop();
                ExportVariables.Printer_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
            }
            else
            {
                foreach (string printer in PrinterSettings.InstalledPrinters)
                {
                    if (!printer.Contains("Microsoft Print to PDF", StringComparison.OrdinalIgnoreCase) && !printer.Contains("XPS", StringComparison.OrdinalIgnoreCase) && !printer.Contains("OneNote", StringComparison.OrdinalIgnoreCase))
                    {
                        ExportVariables.General_export_TotalTests++;
                        string registryPath = @"SYSTEM\CurrentControlSet\Control\Print\Printers\" + printer;
                        string PrinterIP = Registry.LocalMachine.OpenSubKey(registryPath).GetValue("Location").ToString().Split("//").Last().Split(":").First();

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
                                richTextBoxLogs.AppendText("- IP : " + PrinterIP + Environment.NewLine + "- Ping : Echec" + Environment.NewLine);
                            }
                        }
                        else
                        {
                            richTextBoxLogs.AppendText(printer + Environment.NewLine);
                            richTextBoxLogs.AppendText("- IP : Non trouvé" + Environment.NewLine);
                        }
                    }

                    stopwatch.Stop();
                    ExportVariables.Printer_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
                }
            }
        }

        /// <summary>
        /// Event handler for the Start button click event. 
        /// For output formatting, see the ExecutionSequentielle method instead.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
        /// It's here that you can add or remove tests and format the output.
        /// </summary>
        /// <returns></returns>
        async Task ExecutionSequentielle()
        {
            buttonCopier.Enabled = false;
            buttonExportForm.Enabled = false;
            richTextBoxLogs.Clear();

            try
            {
                Stopwatch stopwatch = new();
                stopwatch.Start();
                richTextBoxLogs.AppendText("AccountTester - Rapport de test" + Environment.NewLine);
                richTextBoxLogs.AppendText(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + Environment.NewLine);
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText($"#### Utilisateur :" + Environment.NewLine);
                richTextBoxLogs.AppendText($"- {ExportVariables.General_export_UserName}" + Environment.NewLine);
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText("#### Internet :" + Environment.NewLine);
                await InternetConnexionTest();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText("#### Lecteurs réseaux :" + Environment.NewLine);
                NetworkStorageRightsTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText("#### Version Office :" + Environment.NewLine);
                OfficeVersionTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                if (_WordIsInstalled)
                {
                    richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                    richTextBoxLogs.AppendText("#### Droits Office :" + Environment.NewLine);
                    OfficeWRTesting();
                    richTextBoxLogs.AppendText(Environment.NewLine);
                }

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText("#### Imprimantes :" + Environment.NewLine);
                PrinterTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText("#### Tests terminés." + Environment.NewLine);
                stopwatch.Stop();
                richTextBoxLogs.AppendText("- Total temps écoulé : " + stopwatch.ElapsedMilliseconds + " ms" + Environment.NewLine);
                richTextBoxLogs.AppendText($"- Tests réussi : {ExportVariables.General_export_TotalSuccess}/{ExportVariables.General_export_TotalTests}");
                ExportVariables.General_export_TotalSuccess = 0;
                ExportVariables.General_export_TotalTests = 0;

                // Play a sound when the tests are done
                System.Media.SoundPlayer player = new(@"C:\Windows\Media\Windows Message Nudge.wav");
                player.Play();

                buttonStart.Text = "Restart";

                ExportVariables.General_export_Resume = richTextBoxLogs.Text;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur Execution Sequentielle : " + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// Event handler for the Export button click event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonExport_Click(object sender, EventArgs e)
        {
            ExportForm exportForm = new();
            exportForm.ShowDialog();
        }

        private void ButtonCopier_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBoxLogs.Text);
        }
    }
}
