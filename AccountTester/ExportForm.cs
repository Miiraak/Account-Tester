namespace AccountTester
{
    public partial class ExportForm : Form
    {
        public ExportForm()
        {
            InitializeComponent();
            comboBoxExtension.SelectedIndex = 0;
        }

        private void ButtonSelectPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new()
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxFilePath.Text = folderBrowserDialog.SelectedPath;
            }
        }

        private void ButtonExport_Click(object sender, EventArgs e)
        {
            string fileName;
            string filePath;
            string extension;

            if (textBoxFileName.Text.Trim() == "")
                fileName = $"TestReport_{Environment.UserName}_{DateTime.Now:yyyyMMddHHmmss}";
            else
                fileName = textBoxFileName.Text;

            if (textBoxFilePath.Text == "")
                filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            else
                filePath = textBoxFilePath.Text;

            if (comboBoxExtension.Text == "")
            {
                MessageBox.Show("Please select a file extension.");
                return;
            }
            else
                extension = comboBoxExtension.Text;

            // Exportation du rapport selon l'extension choisie
            switch (extension)
            {
                case ".pdf":
                    ExportToPdf(fileName, filePath);
                    break;
                case ".csv":
                    ExportToCsv(fileName, filePath);
                    break;
                case ".xml":
                    ExportToXml(fileName, filePath);
                    break;
                case ".json":
                    ExportToJson(fileName, filePath);
                    break;
                case ".xlsx":
                    ExportToExcel(fileName, filePath);
                    break;
                case ".txt":
                    ExportToTxt(fileName, filePath);
                    break;
            }
        }

        private void ExportToPdf(string fileName, string filePath)
        {
            // Code pour exporter le rapport en PDF
        }

        private void ExportToCsv(string fileName, string filePath)
        {
            // Code pour exporter le rapport en CSV
        }

        private void ExportToXml(string fileName, string filePath)
        {
            // Code pour exporter le rapport en XML
        }

        private void ExportToJson(string fileName, string filePath)
        {
            // Code pour exporter le rapport en JSON
        }

        private void ExportToExcel(string fileName, string filePath)
        {
            // Code pour exporter le rapport en Excel
        }

        private void ExportToTxt(string fileName, string filePath)
        {
            // Code pour exporter le rapport en TXT
            // Création du fichier
            string path = Path.Combine(filePath, $"{fileName}.txt");
            int i = 0;

            using StreamWriter sw = new(path);
            // Création de l'entête
            sw.WriteLine("Test Report");
            sw.WriteLine($"Date: {DateTime.Today:yyyy-MM-dd}");
            sw.WriteLine();
            sw.WriteLine("Général");
            sw.WriteLine($"Système d'exploitation: {ExportVariables.General_export_DeviceOS}");
            sw.WriteLine($"Type de processus: {ExportVariables.General_export_ProcessType}");
            sw.WriteLine($"Architecture du système d'exploitation: {ExportVariables.General_export_OSArchitecture}");
            sw.WriteLine($"Nom d'utilisateur: {ExportVariables.General_export_UserName}");
            sw.WriteLine();

            // Connexion Internet
            sw.WriteLine("Connexion Internet");
            sw.WriteLine($"Date et heure: {ExportVariables.InternetConnexion_export_DateAndHour}");
            sw.WriteLine($"Type de connexion: {ExportVariables.InternetConnexion_export_ConnexionType}");
            sw.WriteLine($"URL testée: {ExportVariables.InternetConnexion_export_TestedURL}");
            sw.WriteLine($"Statut HTML: {ExportVariables.InternetConnexion_export_HTMLStatut}");
            sw.WriteLine($"Temps de réponse: {ExportVariables.InternetConnexion_export_ResponseTime}");
            sw.WriteLine();

            // Droits de stockage réseau
            sw.WriteLine("Droits de stockage réseau");
            sw.WriteLine($"Date et heure: {ExportVariables.NetworkStorageRights_export_DateAndHour}");
            sw.WriteLine($"Protocole utilisé: {ExportVariables.NetworkStorageRights_export_UsedProtocol}");
            sw.WriteLine($"Type de connexion: {ExportVariables.NetworkStorageRights_export_ConnexionType}");
            sw.WriteLine("Lettre de disque");
            if (ExportVariables.NetworkStorageRights_export_DiskLetter != null)
            {
                foreach (string diskLetter in ExportVariables.NetworkStorageRights_export_DiskLetter)
                {
                    sw.WriteLine($"- {diskLetter}");
                    sw.WriteLine($"Chemin UNC: {ExportVariables.NetworkStorageRights_export_CheminUNC?[i]}");
                    sw.WriteLine($"Serveur: {ExportVariables.NetworkStorageRights_export_Serveur?[i]}");
                    sw.WriteLine($"Nom du partage: {ExportVariables.NetworkStorageRights_export_ShareName?[i]}");
                    i++;
                }
            }
            else
            {
                sw.WriteLine("Aucun disque réseau trouvé");
            }
            i = 0;
            sw.WriteLine();

            // Version d'Office
            sw.WriteLine("Version d'Office");
            sw.WriteLine($"Date et heure: {ExportVariables.OfficeVersion_export_DateAndHour}");
            sw.WriteLine($"Version d'Office: {ExportVariables.OfficeVersion_export_OfficeVersion}");
            sw.WriteLine($"Édition d'Office: {ExportVariables.OfficeVersion_export_OfficeEdition}");
            sw.WriteLine($"Architecture d'Office: {ExportVariables.OfficeVersion_export_OfficeArchitecture}");
            sw.WriteLine($"Chemin d'Office: {ExportVariables.OfficeVersion_export_OfficePath}");
            sw.WriteLine($"ID de produit d'Office: {ExportVariables.OfficeVersion_export_OfficeProductID}");
            sw.WriteLine($"Numéro de série d'Office: {ExportVariables.OfficeVersion_export_OfficeSerialNumber}");
            sw.WriteLine($"Statut du numéro de série d'Office: {ExportVariables.OfficeVersion_export_OfficeSerialNumberStatus}");
            sw.WriteLine();

            // Droits d'Office
            sw.WriteLine("Droits d'Office");
            sw.WriteLine($"Date et heure: {ExportVariables.OfficeRights_export_DateAndHour}");
            sw.WriteLine($"Peut écrire: {ExportVariables.OfficeRights_export_CanWrite}");
            sw.WriteLine($"Peut lire: {ExportVariables.OfficeRights_export_CanRead}");
            sw.WriteLine($"Peut exécuter: {ExportVariables.OfficeRights_export_CanExecute}");
            sw.WriteLine($"Peut supprimer: {ExportVariables.OfficeRights_export_CanDelete}");
            sw.WriteLine($"Peut copier: {ExportVariables.OfficeRights_export_CanCopy}");
            sw.WriteLine($"Peut déplacer: {ExportVariables.OfficeRights_export_CanMove}");
            sw.WriteLine($"Peut renommer: {ExportVariables.OfficeRights_export_CanRename}");
            sw.WriteLine($"Peut créer: {ExportVariables.OfficeRights_export_CanCreate}");
            sw.WriteLine("Dossier testé");
            if (ExportVariables.OfficeRights_export_FolderTested != null)
            {
                foreach (string folderTested in ExportVariables.OfficeRights_export_FolderTested)
                {
                    sw.WriteLine($"- {folderTested}");
                }
            }
            else
            {
                sw.WriteLine("Aucun dossier testé");
            }
            sw.WriteLine();

            // Imprimante 
            sw.WriteLine("Imprimante");
            sw.WriteLine($"Date et heure: {ExportVariables.Printer_export_DateAndHour}");
            sw.WriteLine("Nom de l'imprimante");
            if (ExportVariables.Printer_export_PrinterName != null)
            {
                foreach (string printerName in ExportVariables.Printer_export_PrinterName)
                {
                    sw.WriteLine($"- {printerName}");
                    sw.WriteLine($"Statut de l'imprimante: {ExportVariables.Printer_export_PrinterStatus?[i]}");
                    sw.WriteLine($"Pilote de l'imprimante: {ExportVariables.Printer_export_PrinterDriver?[i]}");
                    sw.WriteLine($"Port de l'imprimante: {ExportVariables.Printer_export_PrinterPort?[i]}");
                    sw.WriteLine($"Emplacement de l'imprimante: {ExportVariables.Printer_export_PrinterLocation?[i]}");
                    i++;
                }
            }
            else
            {
                sw.WriteLine("Aucune imprimante trouvée");
            }
            i = 0;
            sw.WriteLine();

            // Résumé
            sw.WriteLine("Résumé");
            sw.WriteLine(ExportVariables.General_export_Resume);

            // Fermeture du StreamWriter
            sw.Close();

            MessageBox.Show("The report has been exported successfully.");
            this.Close();
        }
    }
}
