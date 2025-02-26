namespace AccountTester
{
    public partial class ExportForm : Form
    {
        public ExportForm()
        {
            InitializeComponent();
            comboBoxExtension.SelectedIndex = 0;
        }

        /// <summary>
        /// This method allows the user to select a path to save the report.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>
        /// Main call to export the report according to the selected extension, name and path.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                case ".csv":
                    break;
                case ".xml":
                    break;
                case ".json":
                    break;
                case ".txt":
                    ExportToTxt(fileName, filePath);
                    break;
                case ".log":
                    ExportToLog(fileName, filePath);
                    break;
            }
        }

        /// <summary>
        /// Export the report to a TXT file.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="filePath"></param>
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
            sw.WriteLine($"Nom d'utilisateur: {ExportVariables.General_export_UserName}");
            sw.WriteLine();
            sw.WriteLine();

            // Général
            sw.WriteLine("Général");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Système d'exploitation: {ExportVariables.General_export_DeviceOS}");
            sw.WriteLine($"Type de processus: {ExportVariables.General_export_ProcessType}");
            sw.WriteLine($"Architecture de l'OS: {ExportVariables.General_export_OSArchitecture}");
            sw.WriteLine();
            sw.WriteLine();

            // Connexion Internet
            sw.WriteLine("Connexion Internet");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Date et heure: {ExportVariables.InternetConnexion_export_DateAndHour}");
            sw.WriteLine($"URL testée: {ExportVariables.InternetConnexion_export_TestedURL}");
            sw.WriteLine($"Statut HTML: {ExportVariables.InternetConnexion_export_HTMLStatut}");
            sw.WriteLine($"Temps de réponse: {ExportVariables.InternetConnexion_export_ElapsedTime} ms");
            sw.WriteLine();
            sw.WriteLine();

            // Droits de stockage réseau
            sw.WriteLine("Droits de stockage réseau");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Date et heure: {ExportVariables.NetworkStorageRights_export_DateAndHour}");
            sw.WriteLine($"Type de connexion: {ExportVariables.NetworkStorageRights_export_ConnexionType}");
            sw.WriteLine("Lettre de disque");
            if (ExportVariables.NetworkStorageRights_export_DiskLetter != null)
            {
                foreach (string diskLetter in ExportVariables.NetworkStorageRights_export_DiskLetter)
                {
                    sw.WriteLine($"Letter: {diskLetter}");
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
            sw.WriteLine($"Temps écoulé : {ExportVariables.NetworkStorageRights_export_ElapsedTime} ms");
            i = 0;
            sw.WriteLine();
            sw.WriteLine();

            // Version d'Office
            sw.WriteLine("Version d'Office");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Date et heure: {ExportVariables.OfficeVersion_export_DateAndHour}");
            sw.WriteLine($"Version d'Office: {ExportVariables.OfficeVersion_export_OfficeVersion}");
            sw.WriteLine($"Édition d'Office: {ExportVariables.OfficeVersion_export_OfficeEdition}");
            sw.WriteLine($"Architecture d'Office: {ExportVariables.OfficeVersion_export_OfficeArchitecture}");
            sw.WriteLine($"Chemin d'Office: {ExportVariables.OfficeVersion_export_OfficePath}");
            sw.WriteLine($"ID de produit d'Office: {ExportVariables.OfficeVersion_export_OfficeProductID}");
            sw.WriteLine($"Numéro de série d'Office: {ExportVariables.OfficeVersion_export_OfficeSerialNumber}");
            sw.WriteLine($"Statut du numéro de série d'Office: {ExportVariables.OfficeVersion_export_OfficeSerialNumberStatus}");
            sw.WriteLine($"Temps écoulé : {ExportVariables.OfficeVersion_export_ElapsedTime} ms");
            sw.WriteLine();
            sw.WriteLine();

            // Droits d'Office
            sw.WriteLine("Droits d'Office");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Date et heure: {ExportVariables.OfficeRights_export_DateAndHour}");
            sw.WriteLine($"Peut écrire: {ExportVariables.OfficeRights_export_CanWrite}");
            sw.WriteLine($"Peut lire: {ExportVariables.OfficeRights_export_CanRead}");
            sw.WriteLine($"Peut supprimer: {ExportVariables.OfficeRights_export_CanDelete}");
            sw.WriteLine($"Peut copier: {ExportVariables.OfficeRights_export_CanCopy}");
            sw.WriteLine($"Peut déplacer: {ExportVariables.OfficeRights_export_CanMove}");
            sw.WriteLine($"Peut renommer: {ExportVariables.OfficeRights_export_CanRename}");
            sw.WriteLine($"Peut créer: {ExportVariables.OfficeRights_export_CanCreate}");
            sw.WriteLine("Dossier testé: ");
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
            sw.WriteLine($"Temps écoulé : {ExportVariables.OfficeRights_export_ElapsedTime} ms");
            sw.WriteLine();
            sw.WriteLine();

            // Imprimante 
            sw.WriteLine("Imprimante");
            sw.WriteLine("-----------------------------------");
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
                    i++;
                }
            }
            else
            {
                sw.WriteLine("Aucune imprimante trouvée");
            }
            sw.WriteLine($"Temps écoulé : {ExportVariables.Printer_export_ElapsedTime} ms");
            i = 0;
            sw.WriteLine();
            sw.WriteLine();

            /*
            // Résumé
            sw.WriteLine("Résumé");
            sw.WriteLine(ExportVariables.General_export_Resume);
            */

            // Fermeture du StreamWriter
            sw.Close();

            MessageBox.Show("The report has been exported successfully.");
            this.Close();
        }

        private void ExportToLog(string fileName, string filePath)
        {
            string path = Path.Combine(filePath, $"{fileName}.log");
            using StreamWriter sw = new(path);

            sw.Write(ExportVariables.General_export_Resume);
            sw.Close();

            MessageBox.Show("The report has been exported successfully.");
            this.Close();
        }
    }
}
