using System.Text.Json;
using System.Xml;

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
                    ExportToCSV(fileName, filePath);
                    break;
                case ".xml":
                    ExportToXML(fileName, filePath);
                    break;
                case ".json":
                    ExportToJSON(fileName, filePath);
                    break;
                case ".txt":
                    ExportToTxt(fileName, filePath);
                    break;
                case ".log":
                    ExportToLog(fileName, filePath);
                    break;
                case ".zip":
                    ExportToZip(fileName, filePath);
                    break;
            }

            MessageBox.Show("The report has been exported successfully.");
            this.Close();
        }

        private void ExportToTxt(string fileName, string filePath)
        {
            // Code pour exporter le rapport en TXT
            // Création du fichier
            string path = Path.Combine(filePath, $"{fileName}.txt");
            int i = 0;

            using StreamWriter sw = new(path);
            // Création de l'entête
            sw.WriteLine(fileName);
            sw.WriteLine($"Date: {ExportVariables.General_DateAndHour}");
            sw.WriteLine($"Username: {ExportVariables.General_export_UserName}\n\n");

            // Général
            sw.WriteLine("General");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Operating system: {ExportVariables.General_export_DeviceOS}");
            sw.WriteLine($"OS architectury: {ExportVariables.General_export_OSArchitecture}\n\n");

            // Connexion Internet
            sw.WriteLine("Internet connection");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Hour: {ExportVariables.InternetConnexion_export_Hour}");
            sw.WriteLine($"Tested URL: {ExportVariables.InternetConnexion_export_TestedURL}");
            sw.WriteLine($"HTML status: {ExportVariables.InternetConnexion_export_HTMLStatut}");
            sw.WriteLine($"Response time: {ExportVariables.InternetConnexion_export_ElapsedTime} ms\n\n");

            // Droits de stockage réseau
            sw.WriteLine("Network storage rights");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Hour: {ExportVariables.NetworkStorageRights_export_Hour}");
            sw.WriteLine($"Connexion Type : {ExportVariables.NetworkStorageRights_export_ConnexionType}");
            sw.WriteLine("Disk:");
            if (ExportVariables.NetworkStorageRights_export_DiskLetter != null)
            {
                foreach (string diskLetter in ExportVariables.NetworkStorageRights_export_DiskLetter)
                {
                    sw.WriteLine($"Letter: {diskLetter}");
                    sw.WriteLine($"UNC path: {ExportVariables.NetworkStorageRights_export_CheminUNC?[i]}");
                    sw.WriteLine($"Server: {ExportVariables.NetworkStorageRights_export_Serveur?[i]}");
                    sw.WriteLine($"Share name: {ExportVariables.NetworkStorageRights_export_ShareName?[i]}");
                    i++;
                }
            }
            else
            {
                sw.WriteLine("No network share found");
            }
            sw.WriteLine($"Time elapsed: {ExportVariables.NetworkStorageRights_export_ElapsedTime} ms\n\n");
            i = 0;

            // Version d'Office
            sw.WriteLine("Office version");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Hour: {ExportVariables.OfficeVersion_export_Hour}");
            sw.WriteLine($"Office version: {ExportVariables.OfficeVersion_export_OfficeVersion}");
            sw.WriteLine($"Office path: {ExportVariables.OfficeVersion_export_OfficePath}");
            sw.WriteLine($"Time elapsed: {ExportVariables.OfficeVersion_export_ElapsedTime} ms\n\n");

            // Droits d'Office
            sw.WriteLine("Office rights");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Hour: {ExportVariables.OfficeRights_export_Hour}");
            sw.WriteLine($"Can write: {ExportVariables.OfficeRights_export_CanWrite}");
            sw.WriteLine($"Can read: {ExportVariables.OfficeRights_export_CanRead}");
            sw.WriteLine($"Can delete: {ExportVariables.OfficeRights_export_CanDelete}");
            sw.WriteLine($"Can create: {ExportVariables.OfficeRights_export_CanCreate}");
            sw.WriteLine($"Can save: {ExportVariables.OfficeRights_export_CanSave}");
            sw.WriteLine("Folder tested: ");
            sw.WriteLine($"- {ExportVariables.OfficeRights_export_FolderTested}");
            sw.WriteLine();
            sw.WriteLine($"Time elapsed: {ExportVariables.OfficeRights_export_ElapsedTime} ms\n\n");

            // Imprimante 
            sw.WriteLine("Printer");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Hour: {ExportVariables.Printer_export_Hour}");
            sw.WriteLine("Printer name:");
            if (ExportVariables.Printer_export_PrinterName != null)
            {
                foreach (string printerName in ExportVariables.Printer_export_PrinterName)
                {
                    sw.WriteLine($"- {printerName}");
                    sw.WriteLine($"Status: {ExportVariables.Printer_export_PrinterStatus?[i]}");
                    sw.WriteLine($"Printer pilots: {ExportVariables.Printer_export_PrinterDriver?[i]}");
                    sw.WriteLine($"Printer port: {ExportVariables.Printer_export_PrinterPort?[i]}");
                    i++;
                }
            }
            else
            {
                sw.WriteLine("No printer found");
            }
            sw.WriteLine($"Time elapsed: {ExportVariables.Printer_export_ElapsedTime} ms\n\n");
            i = 0;

            // Fermeture du StreamWriter
            sw.Close();
        }

        private void ExportToLog(string fileName, string filePath)
        {
            string path = Path.Combine(filePath, $"{fileName}.log");
            using StreamWriter sw = new(path);

            sw.Write(ExportVariables.General_export_Resume);
            sw.Close();
        }

        private void ExportToXML(string fileName, string filePath)
        {
            XmlDocument doc = new();
            XmlElement root = doc.CreateElement(fileName);
            doc.AppendChild(root);
            XmlElement general = doc.CreateElement("General");
            root.AppendChild(general);
            XmlElement os = doc.CreateElement("OperatingSystem");
            os.InnerText = ExportVariables.General_export_DeviceOS;
            general.AppendChild(os);
            XmlElement osArch = doc.CreateElement("OSArchitecture");
            osArch.InnerText = ExportVariables.General_export_OSArchitecture;
            general.AppendChild(osArch);
            XmlElement userName = doc.CreateElement("UserName");
            userName.InnerText = ExportVariables.General_export_UserName;
            general.AppendChild(userName);
            XmlElement date = doc.CreateElement("Date");
            date.InnerText = ExportVariables.General_DateAndHour;
            general.AppendChild(date);
            XmlElement totalTests = doc.CreateElement("TotalTests");
            totalTests.InnerText = ExportVariables.General_export_TotalTests.ToString();
            general.AppendChild(totalTests);
            XmlElement totalSuccess = doc.CreateElement("TotalSuccess");
            totalSuccess.InnerText = ExportVariables.General_export_TotalSuccess.ToString();
            general.AppendChild(totalSuccess);
            XmlElement internetConnection = doc.CreateElement("InternetConnection");
            root.AppendChild(internetConnection);
            XmlElement hour = doc.CreateElement("Hour");
            hour.InnerText = ExportVariables.InternetConnexion_export_Hour;
            internetConnection.AppendChild(hour);
            XmlElement testedURL = doc.CreateElement("TestedURL");
            testedURL.InnerText = ExportVariables.InternetConnexion_export_TestedURL;
            internetConnection.AppendChild(testedURL);
            XmlElement htmlStatus = doc.CreateElement("HTMLStatus");
            htmlStatus.InnerText = ExportVariables.InternetConnexion_export_HTMLStatut;
            internetConnection.AppendChild(htmlStatus);
            XmlElement responseTime = doc.CreateElement("ResponseTime");
            responseTime.InnerText = ExportVariables.InternetConnexion_export_ElapsedTime;
            internetConnection.AppendChild(responseTime);
            XmlElement networkStorageRights = doc.CreateElement("NetworkStorageRights");
            root.AppendChild(networkStorageRights);
            XmlElement connexionType = doc.CreateElement("ConnexionType");
            connexionType.InnerText = ExportVariables.NetworkStorageRights_export_ConnexionType;
            networkStorageRights.AppendChild(connexionType);
            XmlElement diskLetter = doc.CreateElement("DiskLetter");
            if (ExportVariables.NetworkStorageRights_export_DiskLetter != null)
            {
                foreach (string disk in ExportVariables.NetworkStorageRights_export_DiskLetter)
                {
                    XmlElement diskElement = doc.CreateElement("Disk");
                    diskElement.InnerText = disk;
                    diskLetter.AppendChild(diskElement);
                }
            }
            networkStorageRights.AppendChild(diskLetter);
            XmlElement uncPath = doc.CreateElement("UNCPath");
            if (ExportVariables.NetworkStorageRights_export_CheminUNC != null)
            {
                foreach (string unc in ExportVariables.NetworkStorageRights_export_CheminUNC)
                {
                    XmlElement uncElement = doc.CreateElement("Path");
                    uncElement.InnerText = unc;
                    uncPath.AppendChild(uncElement);
                }
            }
            networkStorageRights.AppendChild(uncPath);
            XmlElement server = doc.CreateElement("Server");
            if (ExportVariables.NetworkStorageRights_export_Serveur != null)
            {
                foreach (string srv in ExportVariables.NetworkStorageRights_export_Serveur)
                {
                    XmlElement serverElement = doc.CreateElement("ServerName");
                    serverElement.InnerText = srv;
                    server.AppendChild(serverElement);
                }
            }
            networkStorageRights.AppendChild(server);
            XmlElement shareName = doc.CreateElement("ShareName");
            if (ExportVariables.NetworkStorageRights_export_ShareName != null)
            {
                foreach (string share in ExportVariables.NetworkStorageRights_export_ShareName)
                {
                    XmlElement shareElement = doc.CreateElement("Share");
                    shareElement.InnerText = share;
                    shareName.AppendChild(shareElement);
                }
            }
            networkStorageRights.AppendChild(shareName);
            XmlElement elapsedTime = doc.CreateElement("ElapsedTime");
            elapsedTime.InnerText = ExportVariables.NetworkStorageRights_export_ElapsedTime;
            networkStorageRights.AppendChild(elapsedTime);
            XmlElement officeVersion = doc.CreateElement("OfficeVersion");
            root.AppendChild(officeVersion);
            XmlElement officeHour = doc.CreateElement("Hour");
            officeHour.InnerText = ExportVariables.OfficeVersion_export_Hour;
            officeVersion.AppendChild(officeHour);
            XmlElement officeVersionElement = doc.CreateElement("Version");
            officeVersionElement.InnerText = ExportVariables.OfficeVersion_export_OfficeVersion;
            officeVersion.AppendChild(officeVersionElement);
            XmlElement officePath = doc.CreateElement("Path");
            officePath.InnerText = ExportVariables.OfficeVersion_export_OfficePath;
            officeVersion.AppendChild(officePath);
            XmlElement officeElapsedTime = doc.CreateElement("ElapsedTime");
            officeElapsedTime.InnerText = ExportVariables.OfficeVersion_export_ElapsedTime;
            officeVersion.AppendChild(officeElapsedTime);
            XmlElement officeRights = doc.CreateElement("OfficeRights");
            root.AppendChild(officeRights);
            XmlElement officeRightsHour = doc.CreateElement("Hour");
            officeRightsHour.InnerText = ExportVariables.OfficeRights_export_Hour;
            officeRights.AppendChild(officeRightsHour);
            XmlElement canWrite = doc.CreateElement("CanWrite");
            canWrite.InnerText = ExportVariables.OfficeRights_export_CanWrite;
            officeRights.AppendChild(canWrite);
            XmlElement canRead = doc.CreateElement("CanRead");
            canRead.InnerText = ExportVariables.OfficeRights_export_CanRead;
            officeRights.AppendChild(canRead);
            XmlElement canDelete = doc.CreateElement("CanDelete");
            canDelete.InnerText = ExportVariables.OfficeRights_export_CanDelete;
            officeRights.AppendChild(canDelete);
            XmlElement canCreate = doc.CreateElement("CanCreate");
            canCreate.InnerText = ExportVariables.OfficeRights_export_CanCreate;
            officeRights.AppendChild(canCreate);
            XmlElement canSave = doc.CreateElement("CanSave");
            canSave.InnerText = ExportVariables.OfficeRights_export_CanSave;
            officeRights.AppendChild(canSave);
            XmlElement testedFolder = doc.CreateElement("TestedFolder");
            testedFolder.InnerText = ExportVariables.OfficeRights_export_FolderTested;
            officeRights.AppendChild(testedFolder);
            XmlElement officeRightsElapsedTime = doc.CreateElement("ElapsedTime");
            officeRightsElapsedTime.InnerText = ExportVariables.OfficeRights_export_ElapsedTime;
            officeRights.AppendChild(officeRightsElapsedTime);
            XmlElement printer = doc.CreateElement("Printer");
            root.AppendChild(printer);
            XmlElement printerHour = doc.CreateElement("Hour");
            printerHour.InnerText = ExportVariables.Printer_export_Hour;
            printer.AppendChild(printerHour);
            XmlElement printerName = doc.CreateElement("PrinterName");
            if (ExportVariables.Printer_export_PrinterName != null)
            {
                foreach (string printerNameElement in ExportVariables.Printer_export_PrinterName)
                {
                    XmlElement printerElement = doc.CreateElement("Printer");
                    printerElement.InnerText = printerNameElement;
                    printerName.AppendChild(printerElement);
                }
            }
            printer.AppendChild(printerName);
            XmlElement printerStatus = doc.CreateElement("PrinterStatus");
            if (ExportVariables.Printer_export_PrinterStatus != null)
            {
                foreach (string printerStatusElement in ExportVariables.Printer_export_PrinterStatus)
                {
                    XmlElement statusElement = doc.CreateElement("Status");
                    statusElement.InnerText = printerStatusElement;
                    printerStatus.AppendChild(statusElement);
                }
            }
            printer.AppendChild(printerStatus);
            XmlElement printerDriver = doc.CreateElement("PrinterDriver");
            if (ExportVariables.Printer_export_PrinterDriver != null)
            {
                foreach (string printerDriverElement in ExportVariables.Printer_export_PrinterDriver)
                {
                    XmlElement driverElement = doc.CreateElement("Driver");
                    driverElement.InnerText = printerDriverElement;
                    printerDriver.AppendChild(driverElement);
                }
            }
            printer.AppendChild(printerDriver);
            XmlElement printerPort = doc.CreateElement("PrinterPort");
            if (ExportVariables.Printer_export_PrinterPort != null)
            {
                foreach (string printerPortElement in ExportVariables.Printer_export_PrinterPort)
                {
                    XmlElement portElement = doc.CreateElement("Port");
                    portElement.InnerText = printerPortElement;
                    printerPort.AppendChild(portElement);
                }
            }
            printer.AppendChild(printerPort);
            XmlElement printerElapsedTime = doc.CreateElement("ElapsedTime");
            printerElapsedTime.InnerText = ExportVariables.Printer_export_ElapsedTime;
            printer.AppendChild(printerElapsedTime);
            // Enregistrement du fichier XML
            string path = Path.Combine(filePath, $"{fileName}.xml");
            doc.Save(path);
        }

        private void ExportToCSV(string fileName, string filePath)
        {
            string path = Path.Combine(filePath, $"{fileName}.csv");
            using StreamWriter sw = new(path);
            // Entête - OK
            CSVWL("Section", "Key", "Value", sw);
            // Général - OK
            CSVWL("General", "Report name", fileName, sw);
            CSVWL("General", "Date/Hour", ExportVariables.General_DateAndHour, sw);
            CSVWL("General", "Operating system", ExportVariables.General_export_DeviceOS, sw);
            CSVWL("General", "OS architectury", ExportVariables.General_export_OSArchitecture, sw);
            CSVWL("General", "Username", ExportVariables.General_export_UserName, sw);
            CSVWL("General", "Total tests", ExportVariables.General_export_TotalTests.ToString(), sw);
            CSVWL("General", "Total succes", ExportVariables.General_export_TotalSuccess.ToString(), sw);
            // Connexion Internet - OK
            CSVWL("Internet Connexion", "Hour", ExportVariables.InternetConnexion_export_Hour, sw);
            CSVWL("Internet Connexion", "Tested URL", ExportVariables.InternetConnexion_export_TestedURL, sw);
            CSVWL("Internet Connexion", "HTML status", ExportVariables.InternetConnexion_export_HTMLStatut, sw);
            CSVWL("Internet Connexion", "Response time (ms)", ExportVariables.InternetConnexion_export_ElapsedTime, sw);
            // Droits de stockage réseau - OK
            CSVWL("Network storage rights", "Hour", ExportVariables.NetworkStorageRights_export_Hour, sw);
            CSVWL("Network storage rights", "Connexion type", ExportVariables.NetworkStorageRights_export_ConnexionType, sw);
            if (ExportVariables.NetworkStorageRights_export_DiskLetter != null)
            {
                for (int i = 0; i < ExportVariables.NetworkStorageRights_export_DiskLetter.Length; i++)
                {
                    CSVWL("Network storage rights", "Disk letter", ExportVariables.NetworkStorageRights_export_DiskLetter[i], sw);
                    CSVWL("Network storage rights", "UNC path", ExportVariables.NetworkStorageRights_export_CheminUNC?[i], sw);
                    CSVWL("Network storage rights", "Server", ExportVariables.NetworkStorageRights_export_Serveur?[i], sw);
                    CSVWL("Network storage rights", "Share name", ExportVariables.NetworkStorageRights_export_ShareName?[i], sw);
                }
            }
            else
            {
                CSVWL("Network storage rights", "No network disk found", "", sw);
            }
            CSVWL("Network storage rights", "Time elapsed (ms)", ExportVariables.NetworkStorageRights_export_ElapsedTime, sw);
            // Version d'Office - OK
            CSVWL("Office version", "Hour", ExportVariables.OfficeVersion_export_Hour, sw);
            if (ExportVariables.OfficeVersion_export_OfficeVersion.Split(',').Length > 0)
            {
                foreach (string version in ExportVariables.OfficeVersion_export_OfficeVersion.Split(','))
                {
                    CSVWL("Office version", "Version", version, sw);
                }
            }
            CSVWL("Office version", "Chemin d'Office", ExportVariables.OfficeVersion_export_OfficePath, sw);
            CSVWL("Office version", "Time elapsed (ms)", ExportVariables.OfficeVersion_export_ElapsedTime, sw);
            // Droits d'Office - 
            CSVWL("Office rights", "Hour", ExportVariables.OfficeRights_export_Hour, sw);
            CSVWL("Office rights", "Can write", ExportVariables.OfficeRights_export_CanWrite, sw);
            CSVWL("Office rights", "Can read", ExportVariables.OfficeRights_export_CanRead, sw);
            CSVWL("Office rights", "Can delete", ExportVariables.OfficeRights_export_CanDelete, sw);
            CSVWL("Office rights", "Can create", ExportVariables.OfficeRights_export_CanCreate, sw);
            CSVWL("Office rights", "Can save", ExportVariables.OfficeRights_export_CanSave, sw);
            CSVWL("Office rights", "Tested folder", ExportVariables.OfficeRights_export_FolderTested, sw);
            CSVWL("Office rights", "Time elapsed (ms)", ExportVariables.OfficeRights_export_ElapsedTime, sw);
            // Imprimante
            CSVWL("Printer", "Hour", ExportVariables.Printer_export_Hour, sw);
            if (ExportVariables.Printer_export_PrinterName != null)
            {
                for (int i = 0; i < ExportVariables.Printer_export_PrinterName.Length; i++)
                {
                    CSVWL("Printer", "Printer name", ExportVariables.Printer_export_PrinterName[i], sw);
                    CSVWL("Printer", "Printer status", ExportVariables.Printer_export_PrinterStatus?[i], sw);
                    CSVWL("Printer", "Printer pilots", ExportVariables.Printer_export_PrinterDriver?[i], sw);
                    CSVWL("Printer", "Printer port", ExportVariables.Printer_export_PrinterPort?[i], sw);
                }
            }
            else
            {
                CSVWL("Printer", "No printer found", "", sw);
            }
            CSVWL("Printer", "Time elapsed (ms)", ExportVariables.Printer_export_ElapsedTime, sw);

            sw.Close();
        }

        // ExportToCSV() writeline function
        private static void CSVWL(string value1, string value2, string? value3, StreamWriter sw)
        {
            sw.WriteLine($"{value1},{value2},{value3 ?? string.Empty}");
        }

        private static readonly JsonSerializerOptions CachedJsonSerializerOptions = new() { WriteIndented = true };

        private void ExportToJSON(string fileName, string filePath)
        {
            var Block = new
            {
                Title = fileName,
                Date = ExportVariables.General_DateAndHour,
                Username = ExportVariables.General_export_UserName,
                General = new
                {
                    OperatingSystem = ExportVariables.General_export_DeviceOS,
                    OSArchitecture = ExportVariables.General_export_OSArchitecture,
                    TotalTests = ExportVariables.General_export_TotalTests,
                    TotalSuccess = ExportVariables.General_export_TotalSuccess
                },
                InternetConnection = new
                {
                    Hour = ExportVariables.InternetConnexion_export_Hour,
                    TestedURL = ExportVariables.InternetConnexion_export_TestedURL,
                    HTMLStatus = ExportVariables.InternetConnexion_export_HTMLStatut,
                    ResponseTime = ExportVariables.InternetConnexion_export_ElapsedTime
                },
                NetworkStorageRights = new
                {
                    Hour = ExportVariables.NetworkStorageRights_export_Hour,
                    ConnexionType = ExportVariables.NetworkStorageRights_export_ConnexionType,
                    DiskLetter = ExportVariables.NetworkStorageRights_export_DiskLetter,
                    UNCPath = ExportVariables.NetworkStorageRights_export_CheminUNC,
                    Server = ExportVariables.NetworkStorageRights_export_Serveur,
                    ShareName = ExportVariables.NetworkStorageRights_export_ShareName,
                    ElapsedTime = ExportVariables.NetworkStorageRights_export_ElapsedTime
                },
                OfficeVersion = new
                {
                    Hour = ExportVariables.OfficeVersion_export_Hour,
                    OfficeVersion = ExportVariables.OfficeVersion_export_OfficeVersion,
                    OfficePath = ExportVariables.OfficeVersion_export_OfficePath,
                    ElapsedTime = ExportVariables.OfficeVersion_export_ElapsedTime
                },
                OfficeRights = new
                {
                    Hour = ExportVariables.OfficeRights_export_Hour,
                    CanWrite = ExportVariables.OfficeRights_export_CanWrite,
                    CanRead = ExportVariables.OfficeRights_export_CanRead,
                    CanDelete = ExportVariables.OfficeRights_export_CanDelete,
                    CanCreate = ExportVariables.OfficeRights_export_CanCreate,
                    CanSave = ExportVariables.OfficeRights_export_CanSave,
                    TestedFolder = ExportVariables.OfficeRights_export_FolderTested,
                    ElapsedTime = ExportVariables.OfficeRights_export_ElapsedTime
                },
                Printer = new
                {
                    Hour = ExportVariables.Printer_export_Hour,
                    PrinterName = ExportVariables.Printer_export_PrinterName,
                    PrinterStatus = ExportVariables.Printer_export_PrinterStatus,
                    PrinterDriver = ExportVariables.Printer_export_PrinterDriver,
                    PrinterPort = ExportVariables.Printer_export_PrinterPort,
                    ElapsedTime = ExportVariables.Printer_export_ElapsedTime
                }
            };

            string jsonString = JsonSerializer.Serialize(Block, CachedJsonSerializerOptions);
            File.WriteAllText($"{filePath}\\{fileName}.json", jsonString);
        }

        /// <summary>
        /// Create a zip file that contains all the disponible reports.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="filePath"></param>
        private void ExportToZip(string fileName, string filePath)
        {
            string tempFolder = Path.GetTempPath();

            string tempPath = Path.Combine(tempFolder, fileName);
            Directory.CreateDirectory(tempPath);

            ExportToCSV(fileName, tempPath);
            ExportToXML(fileName, tempPath);
            ExportToJSON(fileName, tempPath);
            ExportToTxt(fileName, tempPath);
            ExportToLog(fileName, tempPath);

            string zipPath = Path.Combine(filePath, $"{fileName}.zip");
            if (File.Exists(zipPath))
            {
                File.Delete(zipPath);
            }
            System.IO.Compression.ZipFile.CreateFromDirectory(tempPath, zipPath);

            Directory.Delete(tempPath, true);
        }
    }
}
