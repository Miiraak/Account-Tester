using Microsoft.Win32;
using System.Configuration;
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
                richTextBoxLogs.AppendText("Internet connection is not OK : " + Environment.NewLine + ex + Environment.NewLine);
            }
        }

        /// <summary>
        /// Method for testing network storage rights on the system by trying to write a file on each network drive.
        /// </summary>
        private void NetworkStorageRightsTesting()
        {
            // Code for testing network storage rights
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

        /// <summary>
        /// Method for testing Office version on the system
        /// </summary>
        private void OfficeVersionTesting()
        { 
            string registryPath = @"SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0";

            using RegistryKey? key = Registry.LocalMachine.OpenSubKey(registryPath);
            string? officeVersion = key?.GetValue("OfficeProductReleaseIds")?.ToString();

            if (!string.IsNullOrEmpty(officeVersion))
            {
                if (officeVersion.Contains(","))
                {
                    foreach (string version in officeVersion.Split(','))
                    {
                        richTextBoxLogs.AppendText($"{version}" + Environment.NewLine);
                    }
                }
                _WordIsInstalled = true;
                return;
            }

            richTextBoxLogs.AppendText("Aucune" + Environment.NewLine);
        }

        /// <summary>
        /// Method for testing Office Write and Read rights on the system by simulating a Word document creation and editing.
        /// </summary>
        private void OfficeWRTesting()
        {
            try
            {
                string filePath = Path.Combine(Path.GetTempPath(), "test_document.docx");

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
                richTextBoxLogs.AppendText($"Erreur : {ex.Message}" + Environment.NewLine);
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
                /*
                // Test d'impression
                try
                {
                    PrintDocument pd = new();
                    pd.PrinterSettings.PrinterName = PrinterSettings.InstalledPrinters[0];
                    pd.PrintPage += (sender, e) =>
                    {
                        e.Graphics.DrawString("Test d'impression", new Font("Arial", 12), new SolidBrush(Color.Black), new PointF(100, 100));
                    };
                    pd.Print();
                    richTextBoxLogs.AppendText("Impression OK" + Environment.NewLine);
                }
                catch (Exception ex)
                {
                    richTextBoxLogs.AppendText("Impression ECHEC : " + ex.Message + Environment.NewLine);
                }     
                */
            }
        }

        private void ButtonStart_Click(object sender, EventArgs e)
        {
            Task.Run(() => ExecutionSequentielle()).Wait();
        }

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
                richTextBoxLogs.AppendText("Erreur : " + ex.Message + Environment.NewLine);
            }
        }
    }
}
