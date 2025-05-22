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

        /// <summary>
        /// Export the report to a TXT file.
        /// </summary>
        /// <param name="fileName">txt file name</param>
        /// <param name="filePath">txt saving path</param>
        private static void ExportToTxt(string fileName, string filePath)
        {
            string path = Path.Combine(filePath, $"{fileName}.txt");
            int i = 0;

            using StreamWriter sw = new(path);
            sw.WriteLine(fileName);
            sw.WriteLine($"Date: {ExportVariables.General_DateAndHour}");
            sw.WriteLine($"Username: {ExportVariables.General_export_UserName}\n\n");

            sw.WriteLine("General");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Operating system: {ExportVariables.General_export_DeviceOS}");
            sw.WriteLine($"OS architectury: {ExportVariables.General_export_OSArchitecture}\n\n");

            sw.WriteLine("Internet connection");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Hour: {ExportVariables.InternetConnexion_export_Hour}");
            sw.WriteLine($"Tested URL: {ExportVariables.InternetConnexion_export_TestedURL}");
            sw.WriteLine($"HTML status: {ExportVariables.InternetConnexion_export_HTMLStatut}");
            sw.WriteLine($"Response time: {ExportVariables.InternetConnexion_export_ElapsedTime} ms\n\n");

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

            sw.WriteLine("Office version");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"Hour: {ExportVariables.OfficeVersion_export_Hour}");
            sw.WriteLine($"Office version: {ExportVariables.OfficeVersion_export_OfficeVersion}");
            sw.WriteLine($"Office path: {ExportVariables.OfficeVersion_export_OfficePath}");
            sw.WriteLine($"Time elapsed: {ExportVariables.OfficeVersion_export_ElapsedTime} ms\n\n");

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

            sw.Close();
        }

        /// <summary>
        /// Export the log to a file.
        /// </summary>
        /// <param name="fileName">log file name</param>
        /// <param name="filePath">log saving path</param>
        private static void ExportToLog(string fileName, string filePath)
        {
            string path = Path.Combine(filePath, $"{fileName}.log");
            using StreamWriter sw = new(path);

            sw.Write(ExportVariables.General_export_Resume);
            sw.Close();
        }

        /// <summary>
        /// Export the report to a XML file.
        /// </summary>
        /// <param name="fileName">xml file name</param>
        /// <param name="filePath">xml saving path</param>
        private static void ExportToXML(string fileName, string filePath)
        {
            XmlDocument doc = new();

            XmlElement root = doc.CreateElement(fileName);
            doc.AppendChild(root);

            XmlElement general = doc.CreateElement("General");
            root.AppendChild(general);
            JSONW(doc, general, "OperatingSystem", ExportVariables.General_export_DeviceOS);
            JSONW(doc, general, "OSArchitecture", ExportVariables.General_export_OSArchitecture);
            JSONW(doc, general, "Username", ExportVariables.General_export_UserName);
            JSONW(doc, general, "Date", ExportVariables.General_DateAndHour);
            JSONW(doc, general, "TotalTests", ExportVariables.General_export_TotalTests.ToString());
            JSONW(doc, general, "TotalSuccess", ExportVariables.General_export_TotalSuccess.ToString());

            XmlElement internetConnection = doc.CreateElement("InternetConnection");
            root.AppendChild(internetConnection);
            JSONW(doc, internetConnection, "Hour", ExportVariables.InternetConnexion_export_Hour);
            JSONW(doc, internetConnection, "TestedURL", ExportVariables.InternetConnexion_export_TestedURL);
            JSONW(doc, internetConnection, "HTMLStatus", ExportVariables.InternetConnexion_export_HTMLStatut);
            JSONW(doc, internetConnection, "ResponseTime", ExportVariables.InternetConnexion_export_ElapsedTime);

            XmlElement networkStorageRights = doc.CreateElement("NetworkStorageRights");
            root.AppendChild(networkStorageRights);
            JSONW(doc, networkStorageRights, "Hour", ExportVariables.NetworkStorageRights_export_Hour);
            JSONW(doc, networkStorageRights, "ConnexionType", ExportVariables.NetworkStorageRights_export_ConnexionType);
            JSONWL(doc, networkStorageRights, "DiskLetter", ExportVariables.NetworkStorageRights_export_DiskLetter, "Disk");
            JSONWL(doc, networkStorageRights, "UNCPath", ExportVariables.NetworkStorageRights_export_CheminUNC, "Path");
            JSONWL(doc, networkStorageRights, "Server", ExportVariables.NetworkStorageRights_export_Serveur, "ServerName");
            JSONWL(doc, networkStorageRights, "Share", ExportVariables.NetworkStorageRights_export_ShareName, "ShareName");
            JSONW(doc, networkStorageRights, "ElapsedTime", ExportVariables.NetworkStorageRights_export_ElapsedTime);

            XmlElement officeVersion = doc.CreateElement("OfficeVersion");
            root.AppendChild(officeVersion);
            JSONW(doc, officeVersion, "Hour", ExportVariables.OfficeVersion_export_Hour);
            JSONW(doc, officeVersion, "Version", ExportVariables.OfficeVersion_export_OfficeVersion);
            JSONW(doc, officeVersion, "Path", ExportVariables.OfficeVersion_export_OfficePath);
            JSONW(doc, officeVersion, "ElapsedTime", ExportVariables.OfficeVersion_export_ElapsedTime);

            XmlElement officeRights = doc.CreateElement("OfficeRights");
            root.AppendChild(officeRights);
            JSONW(doc, officeRights, "Hour", ExportVariables.OfficeRights_export_Hour);
            JSONW(doc, officeRights, "CanWrite", ExportVariables.OfficeRights_export_CanWrite);
            JSONW(doc, officeRights, "CanRead", ExportVariables.OfficeRights_export_CanRead);
            JSONW(doc, officeRights, "CanDelete", ExportVariables.OfficeRights_export_CanDelete);
            JSONW(doc, officeRights, "CanCreate", ExportVariables.OfficeRights_export_CanCreate);
            JSONW(doc, officeRights, "CanSave", ExportVariables.OfficeRights_export_CanSave);
            JSONW(doc, officeRights, "TestedFolder", ExportVariables.OfficeRights_export_FolderTested);
            JSONW(doc, officeRights, "ElapsedTime", ExportVariables.OfficeRights_export_ElapsedTime);

            XmlElement printer = doc.CreateElement("Printer");
            root.AppendChild(printer);
            JSONW(doc, printer, "Hour", ExportVariables.Printer_export_Hour);
            JSONWL(doc, printer, "PrinterName", ExportVariables.Printer_export_PrinterName, "Printer");
            JSONWL(doc, printer, "PrinterStatus", ExportVariables.Printer_export_PrinterStatus, "Status");
            JSONWL(doc, printer, "PrinterDriver", ExportVariables.Printer_export_PrinterDriver, "Driver");
            JSONWL(doc, printer, "PrinterPort", ExportVariables.Printer_export_PrinterPort, "Port");
            JSONW(doc, printer, "ElapsedTime", ExportVariables.Printer_export_ElapsedTime);

            string path = Path.Combine(filePath, $"{fileName}.xml");
            doc.Save(path);
        }

        /// <summary>
        /// Write a value in the XML file.
        /// </summary>
        /// <param name="doc">The XML document</param>
        /// <param name="root">The parent node</param>
        /// <param name="title">Name of the child node</param>
        /// <param name="value">A value</param>
        public static void JSONW(XmlDocument doc, XmlElement root, string title, string value)
        {
            XmlElement child = doc.CreateElement(title);
            child.InnerText = value;
            root.AppendChild(child);
        }

        /// <summary>
        /// Write a list of values in the XML file.
        /// </summary>
        /// <param name="doc">The XML document</param>
        /// <param name="root">The parent node</param>
        /// <param name="title">Name of the child node</param>
        /// <param name="value">A list of values</param>
        /// <param name="child_value">Name of the child node for each value</param>
        private static void JSONWL(XmlDocument doc, XmlElement root, string title, string[]? value, string child_value)
        {
            XmlElement child = doc.CreateElement(title);
            if (value != null)
            {
                foreach (string item in value)
                {
                    XmlElement itemChild = doc.CreateElement(child_value);
                    itemChild.InnerText = item;
                    child.AppendChild(itemChild);
                }
            }
            root.AppendChild(child);
        }

        /// <summary>
        /// Export the report to a CSV file.
        /// </summary>
        /// <param name="fileName">csv file name</param>
        /// <param name="filePath">csv saving path</param>
        private static void ExportToCSV(string fileName, string filePath)
        {
            string path = Path.Combine(filePath, $"{fileName}.csv");
            using StreamWriter sw = new(path);
            CSVWL("Section", "Key", "Value", sw);

            CSVWL("General", "Report name", fileName, sw);
            CSVWL("General", "Date/Hour", ExportVariables.General_DateAndHour, sw);
            CSVWL("General", "Operating system", ExportVariables.General_export_DeviceOS, sw);
            CSVWL("General", "OS architectury", ExportVariables.General_export_OSArchitecture, sw);
            CSVWL("General", "Username", ExportVariables.General_export_UserName, sw);
            CSVWL("General", "Total tests", ExportVariables.General_export_TotalTests.ToString(), sw);
            CSVWL("General", "Total succes", ExportVariables.General_export_TotalSuccess.ToString(), sw);

            CSVWL("Internet Connexion", "Hour", ExportVariables.InternetConnexion_export_Hour, sw);
            CSVWL("Internet Connexion", "Tested URL", ExportVariables.InternetConnexion_export_TestedURL, sw);
            CSVWL("Internet Connexion", "HTML status", ExportVariables.InternetConnexion_export_HTMLStatut, sw);
            CSVWL("Internet Connexion", "Response time (ms)", ExportVariables.InternetConnexion_export_ElapsedTime, sw);

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

            CSVWL("Office rights", "Hour", ExportVariables.OfficeRights_export_Hour, sw);
            CSVWL("Office rights", "Can write", ExportVariables.OfficeRights_export_CanWrite, sw);
            CSVWL("Office rights", "Can read", ExportVariables.OfficeRights_export_CanRead, sw);
            CSVWL("Office rights", "Can delete", ExportVariables.OfficeRights_export_CanDelete, sw);
            CSVWL("Office rights", "Can create", ExportVariables.OfficeRights_export_CanCreate, sw);
            CSVWL("Office rights", "Can save", ExportVariables.OfficeRights_export_CanSave, sw);
            CSVWL("Office rights", "Tested folder", ExportVariables.OfficeRights_export_FolderTested, sw);
            CSVWL("Office rights", "Time elapsed (ms)", ExportVariables.OfficeRights_export_ElapsedTime, sw);

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

        /// <summary>
        /// Write a line in the CSV file with 3 values.
        /// </summary>
        private static void CSVWL(string value1, string value2, string? value3, StreamWriter sw)
        {
            sw.WriteLine($"{value1},{value2},{value3 ?? string.Empty}");
        }

        private static readonly JsonSerializerOptions CachedJsonSerializerOptions = new() { WriteIndented = true };

        /// <summary>
        /// Export the report to a JSON file.
        /// </summary>
        /// <param name="fileName">json file name</param>
        /// <param name="filePath">json saving path</param>
        private static void ExportToJSON(string fileName, string filePath)
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
        /// <param name="fileName">zip file name</param>
        /// <param name="filePath">zip saving path</param>
        private static void ExportToZip(string fileName, string filePath)
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
