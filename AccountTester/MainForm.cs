using Microsoft.Win32;
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
        }

        /// <summary>
        /// Method for testing internet connection
        /// </summary>
        private async Task InternetConnexionTest()
        {
            string URL = "https://google.ch/";
            try
            {
                using HttpClient client = new();
                client.Timeout = TimeSpan.FromSeconds(3);
                HttpResponseMessage response = await client.GetAsync(URL);
                if (response.IsSuccessStatusCode)
                {
                    richTextBoxLogs.AppendText("Connecté" + Environment.NewLine);
                }
                else
                {
                    richTextBoxLogs.AppendText("Status : " + response.StatusCode + Environment.NewLine);
                }
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
            try
            {
                foreach (var drive in DriveInfo.GetDrives())
                {
                    if (drive.DriveType == DriveType.Network)
                    {
                        // Test d'ecriture
                        try
                        {
                            string testFile = Path.Combine(drive.RootDirectory.FullName, "test.txt");
                            File.WriteAllText(testFile, "test");
                            File.Delete(testFile);
                            richTextBoxLogs.AppendText($"{drive.Name} : OK" + Environment.NewLine);
                        }
                        catch (UnauthorizedAccessException)
                        {
                            richTextBoxLogs.AppendText($"{drive.Name} : Ecriture refusée" + Environment.NewLine);
                        }
                        catch (IOException)
                        {
                            richTextBoxLogs.AppendText($"{drive.Name} : Erreur" + Environment.NewLine);
                        }
                    }
                    else
                    {
                        richTextBoxLogs.AppendText($"{drive.Name} : Omis" + Environment.NewLine);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur NetworkStorageRights : " + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// Method for testing Office version on the system
        /// </summary>
        private void OfficeVersionTesting()
        {
            try
            {
                string registryPath = @"SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0";

                using RegistryKey? key = Registry.LocalMachine.OpenSubKey(registryPath);
                string? officeVersion = key?.GetValue("OfficeProductReleaseIds")?.ToString();

                if (!string.IsNullOrEmpty(officeVersion))
                {
                    if (officeVersion.Contains(','))
                    {
                        foreach (string version in officeVersion.Split(','))
                        {
                            richTextBoxLogs.AppendText($"{version}" + Environment.NewLine);
                        }
                    }
                    else
                    {
                        richTextBoxLogs.AppendText($"{officeVersion}" + Environment.NewLine);
                    }
                    _WordIsInstalled = true;
                    return;
                }

                richTextBoxLogs.AppendText("Aucune" + Environment.NewLine);
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
            try
            {
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
                    richTextBoxLogs.AppendText("Création : OK" + Environment.NewLine);
                }
                else
                {
                    richTextBoxLogs.AppendText("Création : ECHEC." + Environment.NewLine);
                    return;
                }

                // Modification et sauvegarde
                doc = wordApp.Documents.Open(filePath);
                doc.Content.Text += "\nAjout de texte pour le test de sauvegarde.";
                doc.Save();
                doc.Close();
                richTextBoxLogs.AppendText("Sauvegarde : OK" + Environment.NewLine);

                // Réouverture du document
                doc = wordApp.Documents.Open(filePath);
                string content = doc.Content.Text;
                doc.Close();

                if (content.Contains("Ajout de texte pour le test de sauvegarde"))
                {
                    richTextBoxLogs.AppendText("Écriture : OK" + Environment.NewLine);
                    richTextBoxLogs.AppendText("Lecture : OK" + Environment.NewLine);
                }
                else
                    richTextBoxLogs.AppendText("Lecture ou écriture : ECHEC" + Environment.NewLine);

                // Fermeture et nettoyage
                wordApp.Quit();
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wordApp);

                // Suppression du fichier
                File.Delete(filePath);
                if (!File.Exists(filePath))
                    richTextBoxLogs.AppendText("Fichier de test supprimé." + Environment.NewLine);
                else
                    richTextBoxLogs.AppendText("Fichier de test non supprimé." + Environment.NewLine);
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
            // Gather all printers on the system
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                richTextBoxLogs.AppendText(printer + Environment.NewLine);
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
            Task.Run(() => ExecutionSequentielle()).Wait();
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
                richTextBoxLogs.AppendText("Internet :" + Environment.NewLine + "------------------------------" + Environment.NewLine);
                await InternetConnexionTest();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("Lecteur réseaux :" + Environment.NewLine + "------------------------------" + Environment.NewLine);
                NetworkStorageRightsTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("Version Office :" + Environment.NewLine + "------------------------------" + Environment.NewLine);
                OfficeVersionTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                if (_WordIsInstalled)
                {
                    richTextBoxLogs.AppendText("Droit :" + Environment.NewLine + "------------------------------" + Environment.NewLine);
                    OfficeWRTesting();
                    richTextBoxLogs.AppendText(Environment.NewLine);
                }

                richTextBoxLogs.AppendText("Imprimantes :" + Environment.NewLine + "------------------------------" + Environment.NewLine);
                PrinterTesting();
                richTextBoxLogs.AppendText(Environment.NewLine);

                richTextBoxLogs.AppendText("Tests terminés." + Environment.NewLine);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur Execution Sequentielle : " + Environment.NewLine + ex.Message);
            }
        }
    }
}
