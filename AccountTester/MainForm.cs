using Microsoft.Win32;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Runtime.InteropServices;
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
                    richTextBoxLogs.AppendText("- �tat : Connect�" + Environment.NewLine);
                    ExportVariables.General_export_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText("- �tat : " + response.StatusCode + Environment.NewLine);
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

            try
            {
                stopwatch.Start();

                foreach (var drive in DriveInfo.GetDrives())
                {
                    ExportVariables.General_export_TotalTests++;

                    if (drive.DriveType == DriveType.Network)
                    {
                        // Test d'ecriture
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
                            richTextBoxLogs.AppendText($@"- {drive.Name}\ : Ecriture refus�e" + Environment.NewLine);
                        }
                        catch (IOException)
                        {
                            richTextBoxLogs.AppendText($@"- {drive.Name}\ : Erreur" + Environment.NewLine);
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
                richTextBoxLogs.AppendText("- Aucune version trouv�e" + Environment.NewLine);
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

                // Cr�ation du document
                Word.Document doc = wordApp.Documents.Add();
                doc.Content.Text = "Ceci est un test d'�criture dans un fichier .docx via Word.";
                doc.SaveAs2(filePath);
                doc.Close();
                if (File.Exists(filePath))
                {
                    richTextBoxLogs.AppendText("- Cr�ation : OK" + Environment.NewLine);
                    ExportVariables.OfficeRights_export_CanCreate = "True";
                    ExportVariables.General_export_TotalSuccess++;
                }
                else
                {
                    richTextBoxLogs.AppendText("- Cr�ation : ECHEC." + Environment.NewLine);
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

                // R�ouverture du document
                doc = wordApp.Documents.Open(filePath);
                if (doc.Content.Text.Contains("Ceci est un test d'�criture dans un fichier .docx via Word."))
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
            ExportVariables.General_export_TotalTests++;
            Stopwatch stopwatch = new();
            stopwatch.Start();

            // Gather all printers on the system
            if (PrinterSettings.InstalledPrinters.Count == 0)
            {
                richTextBoxLogs.AppendText("Aucune imprimante trouv�e." + Environment.NewLine);
                stopwatch.Stop();
                ExportVariables.Printer_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
                return;
            }

            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                richTextBoxLogs.AppendText("- " + printer + Environment.NewLine);
            }

            stopwatch.Stop();
            ExportVariables.Printer_export_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
            ExportVariables.General_export_TotalSuccess++;
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
            richTextBoxLogs.Clear();

            try
            {
                Stopwatch stopwatch = new();
                stopwatch.Start();

                richTextBoxLogs.AppendText($"#### Utilisateur :" + Environment.NewLine);
                richTextBoxLogs.AppendText($"- {ExportVariables.General_export_UserName}" + Environment.NewLine);
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText("#### Internet :" + Environment.NewLine);
                await InternetConnexionTest();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText("#### Lecteurs r�seaux :" + Environment.NewLine);
                NetworkStorageRightsTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText("#### Version Office :" + Environment.NewLine);
                OfficeVersionTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                if (_WordIsInstalled)
                {
                    richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                    richTextBoxLogs.AppendText("#### Droit :" + Environment.NewLine);
                    OfficeWRTesting();
                    richTextBoxLogs.AppendText(Environment.NewLine);
                }

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText("#### Imprimantes :" + Environment.NewLine);
                PrinterTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("----------------------------------------" + Environment.NewLine);
                richTextBoxLogs.AppendText("#### Tests termin�s." + Environment.NewLine);
                stopwatch.Stop();
                richTextBoxLogs.AppendText("- Total temps �coul� : " + stopwatch.ElapsedMilliseconds + " ms" + Environment.NewLine);
                richTextBoxLogs.AppendText($"- Tests r�ussi : {ExportVariables.General_export_TotalSuccess}/{ExportVariables.General_export_TotalTests}");

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
            buttonExportForm.Enabled = false;
            ExportForm exportForm = new();
            exportForm.Show();
        }

        private void ButtonCopier_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBoxLogs.Text);
        }
    }
}
