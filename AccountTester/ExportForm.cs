using System.Text.Json;
using System.Xml;

namespace AccountTester
{
    public partial class ExportForm : Form
    {
        static string T(string key) => LangManager.Instance.Translate(key);
        static string TT(string key) => LangManager.Instance.TrimTranslate(key);

        public ExportForm()
        {
            InitializeComponent();
            comboBoxExtension.SelectedIndex = 0;

            UpdateTexts();
            LangManager.Instance.LanguageChanged += UpdateTexts;
        }

        private void UpdateTexts()
        {
            this.Text = T("Export");
            labelFile.Text = $"{T("FileName")} :";
            labelPath.Text = $"{T("Path")} :";
            labelExtension.Text = $"{T("Extension")} :";
            buttonExport.Text = T("Export");
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
                fileName = $"{T("Report")}_{Environment.UserName}_{DateTime.Now:yyyyMMddHHmmss}";
            else
                fileName = textBoxFileName.Text;

            if (textBoxFilePath.Text == "")
                filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            else
                filePath = textBoxFilePath.Text;

            if (comboBoxExtension.Text == "")
            {
                MessageBox.Show($"{T("ExportForm_ButtonExport_MessageBox_Error")}.");
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

            MessageBox.Show($"{T("ExportForm_ButtonExport_MessageBox_Success")}.");
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
            sw.WriteLine($"{T("Date")}: {ExportVariables.General_DateAndHour}");
            sw.WriteLine($"{T("Username")}: {ExportVariables.General_export_UserName}\n\n");

            sw.WriteLine(T("General"));
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"{T("OperatingSystem")}: {ExportVariables.General_export_DeviceOS}");
            sw.WriteLine($"{T("OSArchitecture")}: {ExportVariables.General_export_OSArchitecture}\n\n");

            sw.WriteLine(T("InternetConnexion"));
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"{T("Hour")}: {ExportVariables.InternetConnexion_export_Hour}");
            sw.WriteLine($"{T("TestedURL")}: {ExportVariables.InternetConnexion_export_TestedURL}");
            sw.WriteLine($"{T("HTMLStatus")}: {ExportVariables.InternetConnexion_export_HTMLStatut}");
            sw.WriteLine($"{T("ResponseTime")}: {ExportVariables.InternetConnexion_export_ElapsedTime} ms\n\n");

            sw.WriteLine(T("NetworkStorageRights"));
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"{T("Hour")}: {ExportVariables.NetworkStorageRights_export_Hour}");
            sw.WriteLine($"{T("ConnexionType")} : {ExportVariables.NetworkStorageRights_export_ConnexionType}");
            sw.WriteLine($"{T("Disk")}:");
            if (ExportVariables.NetworkStorageRights_export_DiskLetter != null)
            {
                foreach (string diskLetter in ExportVariables.NetworkStorageRights_export_DiskLetter)
                {
                    sw.WriteLine($"{T("Letter")}: {diskLetter}");
                    sw.WriteLine($"{T("UNCPath")}: {ExportVariables.NetworkStorageRights_export_CheminUNC?[i]}");
                    sw.WriteLine($"{T("Server")}: {ExportVariables.NetworkStorageRights_export_Serveur?[i]}");
                    sw.WriteLine($"{T("ShareName")}: {ExportVariables.NetworkStorageRights_export_ShareName?[i]}");
                    i++;
                }
            }
            else
            {
                sw.WriteLine(T("NoNetworkShare"));
            }
            sw.WriteLine($"{T("ElapsedTime")}: {ExportVariables.NetworkStorageRights_export_ElapsedTime} ms\n\n");
            i = 0;

            sw.WriteLine("Office");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"{T("Hour")}: {ExportVariables.OfficeVersion_export_Hour}");
            sw.WriteLine($"{T("OfficeVersion")}: {ExportVariables.OfficeVersion_export_OfficeVersion}");
            sw.WriteLine($"{T("OfficePath")}: {ExportVariables.OfficeVersion_export_OfficePath}");
            sw.WriteLine($"{T("ElapsedTime")}: {ExportVariables.OfficeVersion_export_ElapsedTime} ms\n\n");

            sw.WriteLine(T("OfficeRights"));
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"{T("Hour")}: {ExportVariables.OfficeRights_export_Hour}");
            sw.WriteLine($"{T("CanWrite")}: {ExportVariables.OfficeRights_export_CanWrite}");
            sw.WriteLine($"{T("CanRead")}: {ExportVariables.OfficeRights_export_CanRead}");
            sw.WriteLine($"{T("CanDelete")}: {ExportVariables.OfficeRights_export_CanDelete}");
            sw.WriteLine($"{T("CanCreate")}: {ExportVariables.OfficeRights_export_CanCreate}");
            sw.WriteLine($"{T("CanSave")}: {ExportVariables.OfficeRights_export_CanSave}");
            sw.WriteLine($"{T("TestedFolder")}: ");
            sw.WriteLine($"- {ExportVariables.OfficeRights_export_FolderTested}\n");
            sw.WriteLine($"{T("ElapsedTime")}: {ExportVariables.OfficeRights_export_ElapsedTime} ms\n\n");

            sw.WriteLine(T("Printer"));
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"{T("Hour")}: {ExportVariables.Printer_export_Hour}");
            sw.WriteLine($"{T("Name")}:");
            if (ExportVariables.Printer_export_PrinterName != null)
            {
                foreach (string printerName in ExportVariables.Printer_export_PrinterName)
                {
                    sw.WriteLine($"- {printerName}");
                    sw.WriteLine($"{T("Status")}: {ExportVariables.Printer_export_PrinterStatus?[i]}");
                    sw.WriteLine($"{T("Driver")}: {ExportVariables.Printer_export_PrinterDriver?[i]}");
                    sw.WriteLine($"{T("Port")}: {ExportVariables.Printer_export_PrinterPort?[i]}");
                    i++;
                }
            }
            else
            {
                sw.WriteLine(T("NoPrinterFound"));
            }
            sw.WriteLine($"{T("ElapsedTime")}: {ExportVariables.Printer_export_ElapsedTime} ms\n\n");
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

            XmlElement general = doc.CreateElement(TT("General"));
            root.AppendChild(general);
            XMLW(doc, general, TT("OperatingSystem"), ExportVariables.General_export_DeviceOS);
            XMLW(doc, general, TT("OSArchitecture"), ExportVariables.General_export_OSArchitecture);
            XMLW(doc, general, TT("Username"), ExportVariables.General_export_UserName);
            XMLW(doc, general, TT("Date"), ExportVariables.General_DateAndHour);
            XMLW(doc, general, TT("TotalTests"), ExportVariables.General_export_TotalTests.ToString());
            XMLW(doc, general, TT("TotalSuccess"), ExportVariables.General_export_TotalSuccess.ToString());

            XmlElement internetConnection = doc.CreateElement(TT("InternetConnexion"));
            root.AppendChild(internetConnection);
            XMLW(doc, internetConnection, TT("Hour"), ExportVariables.InternetConnexion_export_Hour);
            XMLW(doc, internetConnection, TT("TestedURL"), ExportVariables.InternetConnexion_export_TestedURL);
            XMLW(doc, internetConnection, TT("HTMLStatus"), ExportVariables.InternetConnexion_export_HTMLStatut);
            XMLW(doc, internetConnection, TT("ResponseTime"), ExportVariables.InternetConnexion_export_ElapsedTime);

            XmlElement networkStorageRights = doc.CreateElement(TT("NetworkStorageRights"));
            root.AppendChild(networkStorageRights);
            XMLW(doc, networkStorageRights, TT("Hour"), ExportVariables.NetworkStorageRights_export_Hour);
            XMLW(doc, networkStorageRights, TT("ConnexionType"), ExportVariables.NetworkStorageRights_export_ConnexionType);
            XMLWL(doc, networkStorageRights, TT("DiskLetter"), ExportVariables.NetworkStorageRights_export_DiskLetter, TT("Disk"));
            XMLWL(doc, networkStorageRights, TT("UNCPath"), ExportVariables.NetworkStorageRights_export_CheminUNC, TT("Path"));
            XMLWL(doc, networkStorageRights, TT("Server"), ExportVariables.NetworkStorageRights_export_Serveur, TT("ServerName"));
            XMLWL(doc, networkStorageRights, TT("Share"), ExportVariables.NetworkStorageRights_export_ShareName, TT("ShareName"));
            XMLW(doc, networkStorageRights, TT("ElapsedTime"), ExportVariables.NetworkStorageRights_export_ElapsedTime);

            XmlElement officeVersion = doc.CreateElement(TT("OfficeVersion"));
            root.AppendChild(officeVersion);
            XMLW(doc, officeVersion, TT("Hour"), ExportVariables.OfficeVersion_export_Hour);
            XMLW(doc, officeVersion, TT("Version"), ExportVariables.OfficeVersion_export_OfficeVersion);
            XMLW(doc, officeVersion, TT("Path"), ExportVariables.OfficeVersion_export_OfficePath);
            XMLW(doc, officeVersion, TT("ElapsedTime"), ExportVariables.OfficeVersion_export_ElapsedTime);

            XmlElement officeRights = doc.CreateElement(TT("OfficeRights"));
            root.AppendChild(officeRights);
            XMLW(doc, officeRights, TT("Hour"), ExportVariables.OfficeRights_export_Hour);
            XMLW(doc, officeRights, TT("CanWrite"), ExportVariables.OfficeRights_export_CanWrite);
            XMLW(doc, officeRights, TT("CanRead"), ExportVariables.OfficeRights_export_CanRead);
            XMLW(doc, officeRights, TT("CanDelete"), ExportVariables.OfficeRights_export_CanDelete);
            XMLW(doc, officeRights, TT("CanCreate"), ExportVariables.OfficeRights_export_CanCreate);
            XMLW(doc, officeRights, TT("CanSave"), ExportVariables.OfficeRights_export_CanSave);
            XMLW(doc, officeRights, TT("TestedFolder"), ExportVariables.OfficeRights_export_FolderTested);
            XMLW(doc, officeRights, TT("ElapsedTime"), ExportVariables.OfficeRights_export_ElapsedTime);

            XmlElement printer = doc.CreateElement(TT("Printer"));
            root.AppendChild(printer);
            XMLW(doc, printer, TT("Hour"), ExportVariables.Printer_export_Hour);
            XMLWL(doc, printer, TT("Printer") + TT("Name"), ExportVariables.Printer_export_PrinterName, TT("Printer"));
            XMLWL(doc, printer, TT("Printer") + TT("Status"), ExportVariables.Printer_export_PrinterStatus, TT("Status"));
            XMLWL(doc, printer, TT("Printer") + TT("Driver"), ExportVariables.Printer_export_PrinterDriver, TT("Driver"));
            XMLWL(doc, printer, TT("Printer") + TT("Port"), ExportVariables.Printer_export_PrinterPort, TT("Port"));
            XMLW(doc, printer, TT("ElapsedTime"), ExportVariables.Printer_export_ElapsedTime);

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
        public static void XMLW(XmlDocument doc, XmlElement root, string title, string value)
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
        private static void XMLWL(XmlDocument doc, XmlElement root, string title, string[]? value, string child_value)
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
            CSVWL(T("Section"), T("Key"), T("Value"), sw);

            CSVWL(T("General"), T("ReportName"), fileName, sw);
            CSVWL(T("General"), $"{T("Date")}/{T("Hour")}", ExportVariables.General_DateAndHour, sw);
            CSVWL(T("General"), T("OperatingSystem"), ExportVariables.General_export_DeviceOS, sw);
            CSVWL(T("General"), T("OSArchitecture"), ExportVariables.General_export_OSArchitecture, sw);
            CSVWL(T("General"), T("Username"), ExportVariables.General_export_UserName, sw);
            CSVWL(T("General"), T("TotalTests"), ExportVariables.General_export_TotalTests.ToString(), sw);
            CSVWL(T("General"), T("TotalSuccess"), ExportVariables.General_export_TotalSuccess.ToString(), sw);

            CSVWL(T("InternetConnexion"), T("Hour"), ExportVariables.InternetConnexion_export_Hour, sw);
            CSVWL(T("InternetConnexion"), T("TestedURL"), ExportVariables.InternetConnexion_export_TestedURL, sw);
            CSVWL(T("InternetConnexion"), T("HTMLStatus"), ExportVariables.InternetConnexion_export_HTMLStatut, sw);
            CSVWL(T("InternetConnexion"), $"{T("ResponseTime")} (ms)", ExportVariables.InternetConnexion_export_ElapsedTime, sw);

            CSVWL(T("NetworkStorageRights"), T("Hour"), ExportVariables.NetworkStorageRights_export_Hour, sw);
            CSVWL(T("NetworkStorageRights"), T("ConnexionType"), ExportVariables.NetworkStorageRights_export_ConnexionType, sw);
            if (ExportVariables.NetworkStorageRights_export_DiskLetter != null)
            {
                for (int i = 0; i < ExportVariables.NetworkStorageRights_export_DiskLetter.Length; i++)
                {
                    CSVWL(T("NetworkStorageRights"), T("DiskLetter"), ExportVariables.NetworkStorageRights_export_DiskLetter[i], sw);
                    CSVWL(T("NetworkStorageRights"), T("UNCPath"), ExportVariables.NetworkStorageRights_export_CheminUNC?[i], sw);
                    CSVWL(T("NetworkStorageRights"), T("Server"), ExportVariables.NetworkStorageRights_export_Serveur?[i], sw);
                    CSVWL(T("NetworkStorageRights"), T("ShareName"), ExportVariables.NetworkStorageRights_export_ShareName?[i], sw);
                }
            }
            else
            {
                CSVWL(T("NetworkStorageRights"), T("NoNetworkShare"), "", sw);
            }
            CSVWL(T("NetworkStorageRights"), $"{T("ElapsedTime")} (ms)", ExportVariables.NetworkStorageRights_export_ElapsedTime, sw);

            CSVWL(T("OfficeVersion"), T("Hour"), ExportVariables.OfficeVersion_export_Hour, sw);
            if (ExportVariables.OfficeVersion_export_OfficeVersion.Split(',').Length > 0)
            {
                foreach (string version in ExportVariables.OfficeVersion_export_OfficeVersion.Split(','))
                {
                    CSVWL(T("OfficeVersion"), T("Version"), version, sw);
                }
            }
            CSVWL(T("OfficeVersion"), "Chemin d'Office", ExportVariables.OfficeVersion_export_OfficePath, sw);
            CSVWL(T("OfficeVersion"), "Time elapsed (ms)", ExportVariables.OfficeVersion_export_ElapsedTime, sw);

            CSVWL(T("OfficeRights"), T("Hour"), ExportVariables.OfficeRights_export_Hour, sw);
            CSVWL(T("OfficeRights"), T("CanWrite"), ExportVariables.OfficeRights_export_CanWrite, sw);
            CSVWL(T("OfficeRights"), T("CanRead"), ExportVariables.OfficeRights_export_CanRead, sw);
            CSVWL(T("OfficeRights"), T("CanDelete"), ExportVariables.OfficeRights_export_CanDelete, sw);
            CSVWL(T("OfficeRights"), T("CanCreate"), ExportVariables.OfficeRights_export_CanCreate, sw);
            CSVWL(T("OfficeRights"), T("CanSave"), ExportVariables.OfficeRights_export_CanSave, sw);
            CSVWL(T("OfficeRights"), T("TestedFolder"), ExportVariables.OfficeRights_export_FolderTested, sw);
            CSVWL(T("OfficeRights"), $"{T("ElapsedTime")} (ms)", ExportVariables.OfficeRights_export_ElapsedTime, sw);

            CSVWL(T("Printer"), T("Hour"), ExportVariables.Printer_export_Hour, sw);
            if (ExportVariables.Printer_export_PrinterName != null)
            {
                for (int i = 0; i < ExportVariables.Printer_export_PrinterName.Length; i++)
                {
                    CSVWL(T("Printer"), T("Name"), ExportVariables.Printer_export_PrinterName[i], sw);
                    CSVWL(T("Printer"), T("Status"), ExportVariables.Printer_export_PrinterStatus?[i], sw);
                    CSVWL(T("Printer"), T("Driver"), ExportVariables.Printer_export_PrinterDriver?[i], sw);
                    CSVWL(T("Printer"), T("Port"), ExportVariables.Printer_export_PrinterPort?[i], sw);
                }
            }
            else
            {
                CSVWL(T("Printer"), T("NoPrinterFound"), "", sw);
            }
            CSVWL(T("Printer"), $"{T("ElapsedTime")} (ms)", ExportVariables.Printer_export_ElapsedTime, sw);

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
