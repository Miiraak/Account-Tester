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
            sw.WriteLine($"{T("Username")}: {ExportVariables.General_UserName}\n\n");

            sw.WriteLine(T("General"));
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"{T("OperatingSystem")}: {ExportVariables.General_DeviceOS}");
            sw.WriteLine($"{T("OSArchitecture")}: {ExportVariables.General_OSArchitecture}\n\n");

            sw.WriteLine(T("InternetConnexion"));
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"{T("Hour")}: {ExportVariables.InternetConnexion_Hour}");
            sw.WriteLine($"{T("TestedURL")}: {ExportVariables.InternetConnexion_TestedURL}");
            sw.WriteLine($"{T("HTMLStatus")}: {ExportVariables.InternetConnexion_HTMLStatut}");
            sw.WriteLine($"{T("ResponseTime")}: {ExportVariables.InternetConnexion_ElapsedTime} ms\n\n");

            sw.WriteLine(T("NetworkStorageRights"));
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"{T("Hour")}: {ExportVariables.NetworkStorageRights_Hour}");
            sw.WriteLine($"{T("Disk")}:");
            if (ExportVariables.NetworkStorageRights_DiskLetter != null)
            {
                foreach (string diskLetter in ExportVariables.NetworkStorageRights_DiskLetter)
                {
                    sw.WriteLine($"{T("Letter")}: {diskLetter}");
                    sw.WriteLine($"{T("UNCPath")}: {ExportVariables.NetworkStorageRights_CheminUNC?[i]}");
                    sw.WriteLine($"{T("Server")}: {ExportVariables.NetworkStorageRights_Serveur?[i]}");
                    sw.WriteLine($"{T("ShareName")}: {ExportVariables.NetworkStorageRights_ShareName?[i]}");
                    i++;
                }
            }
            else
            {
                sw.WriteLine(T("NoNetworkShare"));
            }
            sw.WriteLine($"{T("ElapsedTime")}: {ExportVariables.NetworkStorageRights_ElapsedTime} ms\n\n");
            i = 0;

            sw.WriteLine("Office");
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"{T("Hour")}: {ExportVariables.OfficeVersion_Hour}");
            if (ExportVariables.OfficeVersion_OfficeVersion.Split(',').Length > 0)
            {
                sw.WriteLine($"{T("OfficeVersion")}:");
                foreach (string version in ExportVariables.OfficeVersion_OfficeVersion.Split(','))
                {
                    sw.WriteLine($"- {version}");
                }
                sw.WriteLine($"{T("Path")}: {ExportVariables.OfficeVersion_OfficePath}");
                sw.WriteLine($"{T("Culture")}: {ExportVariables.OfficeVersion_OfficeCulture}");
                sw.WriteLine($"{T("ExcludedApps")}: {ExportVariables.OfficeVersion_OfficeExcludedApps}");
                sw.WriteLine($"{T("LastUpdateStatus")}: {ExportVariables.OfficeVersion_OfficeLastUpdateStatus}");
            }
            else
            {
                sw.WriteLine(T("MainForm_RTBL_PrinterTesting_NotFound"));
            }
            sw.WriteLine($"{T("ElapsedTime")}: {ExportVariables.OfficeVersion_ElapsedTime} ms\n\n");

            sw.WriteLine(T("OfficeRights"));
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"{T("Hour")}: {ExportVariables.OfficeRights_Hour}");
            sw.WriteLine($"{T("Write")}: {ExportVariables.OfficeRights_Write}");
            sw.WriteLine($"{T("Read")}: {ExportVariables.OfficeRights_Read}");
            sw.WriteLine($"{T("Delete")}: {ExportVariables.OfficeRights_Delete}");
            sw.WriteLine($"{T("Create")}: {ExportVariables.OfficeRights_Create}");
            sw.WriteLine($"{T("Save")}: {ExportVariables.OfficeRights_Save}");
            sw.WriteLine($"{T("TestedFolder")}: ");
            sw.WriteLine($"- {ExportVariables.OfficeRights_FolderTested}\n");
            sw.WriteLine($"{T("ElapsedTime")}: {ExportVariables.OfficeRights_ElapsedTime} ms\n\n");

            sw.WriteLine(T("Printer"));
            sw.WriteLine("-----------------------------------");
            sw.WriteLine($"{T("Hour")}: {ExportVariables.Printer_Hour}");
            if (ExportVariables.Printer_PrinterName != null)
            {
                foreach (string printerName in ExportVariables.Printer_PrinterName)
                {
                    sw.WriteLine($"{T("Name")}: {printerName}");
                    sw.WriteLine($"- {T("IP")}: {ExportVariables.Printer_PrinterIP?[i]}");
                    sw.WriteLine($"- {T("Status")}: {ExportVariables.Printer_PrinterStatus?[i]}");
                    sw.WriteLine($"- {T("Driver")}: {ExportVariables.Printer_PrinterDriver?[i]}");
                    sw.WriteLine($"- {T("Port")}: {ExportVariables.Printer_PrinterPort?[i]}");
                    i++;
                }
            }
            else
            {
                sw.WriteLine(T("NoPrinterFound"));
            }
            sw.WriteLine($"{T("ElapsedTime")}: {ExportVariables.Printer_ElapsedTime} ms\n\n");
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

            sw.Write(ExportVariables.General_Resume);
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
            XMLW(doc, general, TT("OperatingSystem"), ExportVariables.General_DeviceOS);
            XMLW(doc, general, TT("OSArchitecture"), ExportVariables.General_OSArchitecture);
            XMLW(doc, general, TT("Username"), ExportVariables.General_UserName);
            XMLW(doc, general, TT("Date"), ExportVariables.General_DateAndHour);
            XMLW(doc, general, TT("TotalTests"), ExportVariables.General_TotalTests.ToString());
            XMLW(doc, general, TT("TotalSuccess"), ExportVariables.General_TotalSuccess.ToString());

            XmlElement internetConnection = doc.CreateElement(TT("InternetConnexion"));
            root.AppendChild(internetConnection);
            XMLW(doc, internetConnection, TT("Hour"), ExportVariables.InternetConnexion_Hour);
            XMLW(doc, internetConnection, TT("TestedURL"), ExportVariables.InternetConnexion_TestedURL);
            XMLW(doc, internetConnection, TT("HTMLStatus"), ExportVariables.InternetConnexion_HTMLStatut);
            XMLW(doc, internetConnection, TT("ResponseTime"), ExportVariables.InternetConnexion_ElapsedTime);

            XmlElement networkStorageRights = doc.CreateElement(TT("NetworkStorageRights"));
            root.AppendChild(networkStorageRights);
            XMLW(doc, networkStorageRights, TT("Hour"), ExportVariables.NetworkStorageRights_Hour);
            for (int i = 0; i < ExportVariables.NetworkStorageRights_DiskLetter?.Length; i++)
            {
                XmlElement disk = doc.CreateElement($"{TT("Disk")}_{i}");
                networkStorageRights.AppendChild(disk);
                XMLW(doc, disk, TT("DiskLetter"), ExportVariables.NetworkStorageRights_DiskLetter[i]);
                XMLW(doc, disk, TT("UNCPath"), ExportVariables.NetworkStorageRights_CheminUNC?[i] ?? string.Empty);
                XMLW(doc, disk, TT("Server"), ExportVariables.NetworkStorageRights_Serveur?[i] ?? string.Empty);
                XMLW(doc, disk, TT("ShareName"), ExportVariables.NetworkStorageRights_ShareName?[i] ?? string.Empty);
            }

            XMLW(doc, networkStorageRights, TT("ElapsedTime"), ExportVariables.NetworkStorageRights_ElapsedTime);

            XmlElement officeVersion = doc.CreateElement(TT("OfficeVersion"));
            root.AppendChild(officeVersion);
            XMLW(doc, officeVersion, TT("Hour"), ExportVariables.OfficeVersion_Hour);
            XMLW(doc, officeVersion, TT("Version"), ExportVariables.OfficeVersion_OfficeVersion);
            XMLW(doc, officeVersion, TT("Path"), ExportVariables.OfficeVersion_OfficePath);
            XMLW(doc, officeVersion, TT("Culture"), ExportVariables.OfficeVersion_OfficeCulture);
            XMLW(doc, officeVersion, TT("ExcludedApps"), ExportVariables.OfficeVersion_OfficeExcludedApps);
            XMLW(doc, officeVersion, TT("LastUpdateStatus"), ExportVariables.OfficeVersion_OfficeLastUpdateStatus);
            XMLW(doc, officeVersion, TT("ElapsedTime"), ExportVariables.OfficeVersion_ElapsedTime);

            XmlElement officeRights = doc.CreateElement(TT("OfficeRights"));
            root.AppendChild(officeRights);
            XMLW(doc, officeRights, TT("Hour"), ExportVariables.OfficeRights_Hour);
            XMLW(doc, officeRights, TT("Write"), ExportVariables.OfficeRights_Write);
            XMLW(doc, officeRights, TT("Read"), ExportVariables.OfficeRights_Read);
            XMLW(doc, officeRights, TT("Delete"), ExportVariables.OfficeRights_Delete);
            XMLW(doc, officeRights, TT("Create"), ExportVariables.OfficeRights_Create);
            XMLW(doc, officeRights, TT("Save"), ExportVariables.OfficeRights_Save);
            XMLW(doc, officeRights, TT("TestedFolder"), ExportVariables.OfficeRights_FolderTested);
            XMLW(doc, officeRights, TT("ElapsedTime"), ExportVariables.OfficeRights_ElapsedTime);

            XmlElement printers = doc.CreateElement(TT("Printer"));
            root.AppendChild(printers);
            XMLW(doc, printers, TT("Hour"), ExportVariables.Printer_Hour);

            for (int i = 0; i < ExportVariables.Printer_PrinterName?.Length; i++)
            {
                XmlElement printer = doc.CreateElement($"{TT("Printer")}_{i}");
                printers.AppendChild(printer);
                XMLW(doc, printer, TT("Name"), ExportVariables.Printer_PrinterName[i]);
                XMLW(doc, printer, TT("IP"), ExportVariables.Printer_PrinterIP[i]);
                XMLW(doc, printer, TT("Status"), ExportVariables.Printer_PrinterStatus[i]);
                XMLW(doc, printer, TT("Driver"), ExportVariables.Printer_PrinterDriver[i]);
                XMLW(doc, printer, TT("Port"), ExportVariables.Printer_PrinterPort[i]);
            }

            XMLW(doc, printers, TT("ElapsedTime"), ExportVariables.Printer_ElapsedTime);

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
            CSVWL(T("General"), T("OperatingSystem"), ExportVariables.General_DeviceOS, sw);
            CSVWL(T("General"), T("OSArchitecture"), ExportVariables.General_OSArchitecture, sw);
            CSVWL(T("General"), T("Username"), ExportVariables.General_UserName, sw);
            CSVWL(T("General"), T("TotalTests"), ExportVariables.General_TotalTests.ToString(), sw);
            CSVWL(T("General"), T("TotalSuccess"), ExportVariables.General_TotalSuccess.ToString(), sw);

            CSVWL(T("InternetConnexion"), T("Hour"), ExportVariables.InternetConnexion_Hour, sw);
            CSVWL(T("InternetConnexion"), T("TestedURL"), ExportVariables.InternetConnexion_TestedURL, sw);
            CSVWL(T("InternetConnexion"), T("HTMLStatus"), ExportVariables.InternetConnexion_HTMLStatut, sw);
            CSVWL(T("InternetConnexion"), $"{T("ResponseTime")} (ms)", ExportVariables.InternetConnexion_ElapsedTime, sw);
            
            CSVWL(T("NetworkStorageRights"), T("Hour"), ExportVariables.NetworkStorageRights_Hour, sw);
            if (ExportVariables.NetworkStorageRights_DiskLetter != null)
            {
                for (int i = 0; i < ExportVariables.NetworkStorageRights_DiskLetter.Length; i++)
                {
                    CSVWL($"{T("Disk")}_{i}", T("DiskLetter"), ExportVariables.NetworkStorageRights_DiskLetter[i], sw);
                    CSVWL($"{T("Disk")}_{i}", T("UNCPath"), ExportVariables.NetworkStorageRights_CheminUNC[i], sw);
                    CSVWL($"{T("Disk")}_{i}", T("Server"), ExportVariables.NetworkStorageRights_Serveur[i], sw);
                    CSVWL($"{T("Disk")}_{i}", T("ShareName"), ExportVariables.NetworkStorageRights_ShareName[i], sw);
                }
            }
            else
            {
                CSVWL(T("NetworkStorageRights"), T("NoNetworkShare"), "", sw);
            }
            CSVWL(T("NetworkStorageRights"), $"{T("ElapsedTime")} (ms)", ExportVariables.NetworkStorageRights_ElapsedTime, sw);

            CSVWL(T("OfficeVersion"), T("Hour"), ExportVariables.OfficeVersion_Hour, sw);
            if (ExportVariables.OfficeVersion_OfficeVersion.Split(',').Length > 0)
            {
                foreach (string version in ExportVariables.OfficeVersion_OfficeVersion.Split(','))
                {
                    CSVWL(T("OfficeVersion"), T("Version"), version, sw);
                }
            }
            CSVWL(T("OfficeVersion"), T("OfficePath"), ExportVariables.OfficeVersion_OfficePath, sw);
            CSVWL(T("OfficeVersion"), T("Culture"), ExportVariables.OfficeVersion_OfficeCulture, sw);
            CSVWL(T("OfficeVersion"), T("ExcludedApps"), ExportVariables.OfficeVersion_OfficeExcludedApps.Replace(",", " "), sw);
            CSVWL(T("OfficeVersion"), T("LastUpdateStatus"), ExportVariables.OfficeVersion_OfficeLastUpdateStatus, sw);
            CSVWL(T("OfficeVersion"), $"{T("ElapsedTime")} (ms)", ExportVariables.OfficeVersion_ElapsedTime, sw);

            CSVWL(T("OfficeRights"), T("Hour"), ExportVariables.OfficeRights_Hour, sw);
            CSVWL(T("OfficeRights"), T("Write"), ExportVariables.OfficeRights_Write, sw);
            CSVWL(T("OfficeRights"), T("Read"), ExportVariables.OfficeRights_Read, sw);
            CSVWL(T("OfficeRights"), T("Delete"), ExportVariables.OfficeRights_Delete, sw);
            CSVWL(T("OfficeRights"), T("Create"), ExportVariables.OfficeRights_Create, sw);
            CSVWL(T("OfficeRights"), T("Save"), ExportVariables.OfficeRights_Save, sw);
            CSVWL(T("OfficeRights"), T("TestedFolder"), ExportVariables.OfficeRights_FolderTested, sw);
            CSVWL(T("OfficeRights"), $"{T("ElapsedTime")} (ms)", ExportVariables.OfficeRights_ElapsedTime, sw);

            CSVWL(T("Printer"), T("Hour"), ExportVariables.Printer_Hour, sw);
            if (ExportVariables.Printer_PrinterName != null)
            {
                for (int i = 0; i < ExportVariables.Printer_PrinterName.Length; i++)
                {
                    CSVWL($"{T("Printer")}_{i}", T("Name"), ExportVariables.Printer_PrinterName[i], sw);
                    CSVWL($"{T("Printer")}_{i}", T("Status"), ExportVariables.Printer_PrinterStatus[i], sw);
                    CSVWL($"{T("Printer")}_{i}", T("IP"), ExportVariables.Printer_PrinterIP[i], sw);
                    CSVWL($"{T("Printer")}_{i}", T("Driver"), ExportVariables.Printer_PrinterDriver[i], sw);
                    CSVWL($"{T("Printer")}_{i}", T("Port"), ExportVariables.Printer_PrinterPort[i], sw);
                }
            }
            else
            {
                CSVWL(T("Printer"), T("NoPrinterFound"), "", sw);
            }
            CSVWL(T("Printer"), $"{T("ElapsedTime")} (ms)", ExportVariables.Printer_ElapsedTime, sw);

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
            var Printers = new Dictionary<string, object>
            {
                [TT("Hour")] = ExportVariables.Printer_Hour,
            };

            int printerCount = ExportVariables.Printer_PrinterName?.Length ?? 0;
            for (int i = 0; i < printerCount; i++)
            {
                var printer = new Dictionary<string, object>
                {
                    [TT("Name")] = ExportVariables.Printer_PrinterName[i],
                    [TT("IP")] = ExportVariables.Printer_PrinterIP[i],
                    [TT("Status")] = ExportVariables.Printer_PrinterStatus[i],
                    [TT("Driver")] = ExportVariables.Printer_PrinterDriver[i],
                    [TT("Port")] = ExportVariables.Printer_PrinterPort[i]
                };

                Printers[$"{TT("Printer")}_{i + 1}"] = printer;
            }
            Printers[TT("ElapsedTime")] = ExportVariables.Printer_ElapsedTime;

            var OfficeRights = new Dictionary<string, object>
            {
                [TT("Hour")] = ExportVariables.OfficeRights_Hour,
                [TT("Write")] = ExportVariables.OfficeRights_Write,
                [TT("Read")] = ExportVariables.OfficeRights_Read,
                [TT("Delete")] = ExportVariables.OfficeRights_Delete,
                [TT("Create")] = ExportVariables.OfficeRights_Create,
                [TT("Save")] = ExportVariables.OfficeRights_Save,
                [TT("TestedFolder")] = ExportVariables.OfficeRights_FolderTested,
                [TT("ElapsedTime")] = ExportVariables.OfficeRights_ElapsedTime
            };

            var OfficeVersion = new Dictionary<string, object>
            {
                [TT("Hour")] = ExportVariables.OfficeVersion_Hour,
                [TT("OfficeVersion")] = ExportVariables.OfficeVersion_OfficeVersion,
                [TT("Path")] = ExportVariables.OfficeVersion_OfficePath,
                [TT("Culture")] = ExportVariables.OfficeVersion_OfficeCulture,
                [TT("ExcludedApps")] = ExportVariables.OfficeVersion_OfficeExcludedApps,
                [TT("LastUpdateStatus")] = ExportVariables.OfficeVersion_OfficeLastUpdateStatus,
                [TT("ElapsedTime")] = ExportVariables.OfficeVersion_ElapsedTime
            };

            var NetworkStorageRights = new Dictionary<string, object>
            {
                [TT("Hour")] = ExportVariables.NetworkStorageRights_Hour,
            };

            int drivesCount = ExportVariables.NetworkStorageRights_DiskLetter?.Length ?? 0;
            for (int i = 0; i < drivesCount; i++)
            {
                var Disk = new Dictionary<string, object>
                {
                    [TT("DiskLetter")] = ExportVariables.NetworkStorageRights_DiskLetter[i],
                    [TT("UNCPath")] = ExportVariables.NetworkStorageRights_CheminUNC[i],
                    [T("Server")] = ExportVariables.NetworkStorageRights_Serveur[i],
                    [TT("ShareName")] = ExportVariables.NetworkStorageRights_ShareName[i]
                };

                NetworkStorageRights[$"{TT("Disk")}_{i + 1}"] = Disk;
            }
            NetworkStorageRights[TT("ElapsedTime")] = ExportVariables.NetworkStorageRights_ElapsedTime;

            var InternetConnexion = new Dictionary<string, object>
            {
                [TT("Hour")] = ExportVariables.InternetConnexion_Hour,
                [TT("TestedURL")] = ExportVariables.InternetConnexion_TestedURL,
                [TT("HTMLStatus")] = ExportVariables.InternetConnexion_HTMLStatut,
                [TT("ResponseTime")] = ExportVariables.InternetConnexion_ElapsedTime
            };

            var General = new Dictionary<string, object>
            {
                [TT("OperatingSystem")] = ExportVariables.General_DeviceOS,
                [TT("OSArchitecture")] = ExportVariables.General_OSArchitecture,
                [TT("TotalTests")] = ExportVariables.General_TotalTests,
                [TT("TotalSuccess")] = ExportVariables.General_TotalSuccess,
            };

            var block = new Dictionary<string, object>
            {
                [TT("Title")] = fileName,
                [TT("Date")] = ExportVariables.General_DateAndHour,
                [TT("Username")] = ExportVariables.General_UserName,
                [TT("General")] = General,
                [TT("InternetConnexion")] = InternetConnexion,
                [TT("NetworkStorageRights")] = NetworkStorageRights,
                [TT("OfficeVersion")] = OfficeVersion,
                [TT("OfficeRights")] = OfficeRights,
                [TT("Printer")] = Printers
            };

            string jsonString = JsonSerializer.Serialize(block, CachedJsonSerializerOptions);
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
