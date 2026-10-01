using Microsoft.Win32;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Management;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace AccountTester
{
    internal class Tests
    {
        static string T(string key) => LangManager.Instance.Translate(key);

        static readonly Stopwatch stopwatch = new();

        /// <summary>
        /// Tests the internet connection by sending an HTTP GET request to a predefined URL.
        /// </summary>
        /// <remarks>This method performs an asynchronous HTTP GET request to the URL specified in
        /// <c>Variables.Target</c>. It logs the connection status and response details to the
        /// provided <see cref="RichTextBox"/>. The method updates several global variables, including the total number
        /// of tests, the elapsed time for the test, and the HTTP status code of the response.</remarks>
        /// <param name="rtb">The <see cref="RichTextBox"/> used to display logs related to the connection test.</param>
        /// <returns></returns>
        internal static async Task InternetConnexionTest(RichTextBox rtb)
        {
            Variables.General_TotalTests++;
            Variables.InternetConnexion_Hour = DateTime.Now.ToString("HH:mm:ss");

            try
            {
                stopwatch.Restart();

                using HttpClient client = new();
                client.Timeout = TimeSpan.FromMilliseconds(Variables.Timeout);
                string customUserAgent = $"AccountTester/{Variables.Version} ({Environment.OSVersion})";
                client.DefaultRequestHeaders.Add("User-Agent", customUserAgent);
                using HttpResponseMessage response = await client.GetAsync(Variables.Target);

                Variables.InternetConnexion_HTMLStatut = response.StatusCode.ToString();

                if (response.IsSuccessStatusCode)
                {
                    rtb.AppendText($"{T("MainForm_RTBL_Internet_Connected")}{Environment.NewLine}");
                    Variables.General_TotalSuccess++;
                }
                else
                {
                    rtb.AppendText($"{T("MainForm_RTBL_Internet_Others")} {response.StatusCode}{Environment.NewLine}");
                }
            }
            catch (Exception ex)
            {
                rtb.AppendText($"{T("MainForm_RTBL_Internet_Others")} {ex.InnerException?.Message ?? ex.Message}{Environment.NewLine}");
            }

            stopwatch.Stop();
            Variables.InternetConnexion_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>
        /// Tests the read and write access rights for network storage drives and logs the results.
        /// </summary>
        /// <remarks>This method iterates through all available drives on the system, identifies network
        /// drives, and attempts to create, verify, and delete a test file on each network drive to determine access
        /// rights. Results are logged to the provided <see cref="RichTextBox"/> control.</remarks>
        /// <param name="rtb">The <see cref="RichTextBox"/> control where the results of the network storage rights testing will be
        /// appended.</param>
        internal static void NetworkStorageRightsTesting(RichTextBox rtb)
        {
            Variables.NetworkStorageRights_Hour = DateTime.Now.ToString("HH:mm:ss");

            try
            {
                stopwatch.Restart();
                string[] foundDrives = [];

                foreach (var drive in DriveInfo.GetDrives())
                {
                    foundDrives = [.. foundDrives, drive.Name[0].ToString()];
                    Variables.General_TotalTests++;
                    Variables.NetworkStorageRights_DiskLetter = [.. Variables.NetworkStorageRights_DiskLetter, drive.Name];

                    if (drive.DriveType == DriveType.Network)
                    {
                        string cheminUNC = drive.RootDirectory.FullName;
                        string serveur = "";
                        string shareName = "";

                        var uncParts = cheminUNC.TrimEnd('\\').Split('\\');
                        if (uncParts.Length >= 4)
                        {
                            serveur = uncParts[2];
                            shareName = uncParts[3];
                        }
                        else
                        {
                            serveur = T("Unknown");
                            shareName = T("Unknown");
                        }

                        try
                        {
                            string testFile = Path.Combine(drive.RootDirectory.FullName, "test.txt");
                            File.WriteAllText(testFile, "test");

                            if (File.Exists(testFile))
                            {
                                rtb.AppendText($@"- {drive.Name} : OK" + Environment.NewLine);
                                Variables.General_TotalSuccess++;
                                Variables.NetworkStorageRights_CheminUNC = [.. Variables.NetworkStorageRights_CheminUNC, cheminUNC];
                                Variables.NetworkStorageRights_Serveur = [.. Variables.NetworkStorageRights_Serveur, serveur];
                                Variables.NetworkStorageRights_ShareName = [.. Variables.NetworkStorageRights_ShareName, shareName];
                            }

                            File.Delete(testFile);
                        }
                        catch (UnauthorizedAccessException)
                        {
                            rtb.AppendText($@"- {drive.Name} : {T("MainForm_RTBL_NetworkStorageRightsTesting_Refused")}" + Environment.NewLine);
                            Variables.NetworkStorageRights_CheminUNC = [.. Variables.NetworkStorageRights_CheminUNC, T("UnauthorizedAccess")];
                            Variables.NetworkStorageRights_Serveur = [.. Variables.NetworkStorageRights_Serveur, T("UnauthorizedAccess")];
                            Variables.NetworkStorageRights_ShareName = [.. Variables.NetworkStorageRights_ShareName, T("UnauthorizedAccess")];
                        }
                        catch (IOException)
                        {
                            rtb.AppendText($@"- {drive.Name} : {T("MainForm_RTBL_NetworkStorageRightsTesting_Error")}" + Environment.NewLine);
                            Variables.NetworkStorageRights_CheminUNC = [.. Variables.NetworkStorageRights_CheminUNC, T("IOError")];
                            Variables.NetworkStorageRights_Serveur = [.. Variables.NetworkStorageRights_Serveur, T("IOError")];
                            Variables.NetworkStorageRights_ShareName = [.. Variables.NetworkStorageRights_ShareName, T("IOError")];
                        }
                    }
                    else
                    {
                        Variables.NetworkStorageRights_CheminUNC = [.. Variables.NetworkStorageRights_CheminUNC, drive.Name];
                        Variables.NetworkStorageRights_Serveur = [.. Variables.NetworkStorageRights_Serveur, "localhost"];
                        Variables.NetworkStorageRights_ShareName = [.. Variables.NetworkStorageRights_ShareName, T("None")];

                        rtb.AppendText($@"- {drive.Name} : {T("Omitted")}" + Environment.NewLine);
                        Variables.General_TotalSuccess++;
                    }
                }

                string[] drivesList = [.. Variables.DrivesList.Split(';').Select(p => p.Trim()).Where(p => !string.IsNullOrEmpty(p))];
                foreach (string drive in drivesList)
                {
                    if (!foundDrives.Contains(drive))
                    {
                        rtb.AppendText($"- {drive}:\\ : {T("Missing")}" + Environment.NewLine);
                        Variables.General_TotalTests++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error NetworkStorageRights : " + Environment.NewLine + ex);
            }

            stopwatch.Stop();
            Variables.NetworkStorageRights_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>
        /// Tests and retrieves information about the installed version of Microsoft Office.
        /// </summary>
        /// <remarks>This method queries the Windows Registry to gather details about the installed Office
        /// version,  including the product release IDs, installation path, culture, excluded applications, and last
        /// update status.  The results are displayed in the provided <see cref="RichTextBox"/> and stored in global
        /// variables for further use.</remarks>
        /// <param name="rtb">The <see cref="RichTextBox"/> control where the retrieved Office version details will be displayed.</param>
        internal static void OfficeVersionTesting(RichTextBox rtb)
        {
            Variables.General_TotalTests++;
            Variables.OfficeVersion_Hour = DateTime.Now.ToString("HH:mm:ss");

            try
            {
                stopwatch.Restart();

                using RegistryKey? key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0");
                string? officeVersion = key?.GetValue("OfficeProductReleaseIds")?.ToString();

                if (!string.IsNullOrEmpty(officeVersion))
                {
                    Variables.OfficeVersion_OfficeVersion = officeVersion;

                    if (officeVersion.Contains(','))
                    {
                        foreach (string version in officeVersion.Split(','))
                        {
                            rtb.AppendText($"- {version}" + Environment.NewLine);
                        }
                    }
                    else
                    {
                        rtb.AppendText($"- {officeVersion}" + Environment.NewLine);
                    }
                    Variables.General_TotalSuccess++;
                    Variables.WordIsInstalled = true;
                }
                else
                {
                    rtb.AppendText($"- {T("MainForm_RTBL_OfficeVersionTesting_NotFound")}" + Environment.NewLine);
                }

                Variables.OfficeVersion_OfficePath = GetRegValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", "InstallationPath");
                Variables.OfficeVersion_OfficeCulture = GetRegValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0", "OfficeCulture");
                Variables.OfficeVersion_OfficeExcludedApps = GetRegValue(@"SOFTWARE\Microsoft\Office\ClickToRun\Inventory\Office\16.0", "OfficeExcludedApps");
                Variables.OfficeVersion_OfficeLastUpdateStatus = GetRegValue(@"SOFTWARE\Microsoft\Office\ClickToRun\UpdateStatus", "LastUpdateResult");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error OfficeVersion : " + Environment.NewLine + ex.Message);
            }

            stopwatch.Stop();
            Variables.OfficeVersion_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
        }

        /// <summary>
        /// Retrieves the specified value from a registry key located in the Local Machine hive.
        /// </summary>
        /// <remarks>This method accesses the registry key in the Local Machine hive. Ensure the
        /// application has appropriate permissions to read from the registry. If the specified value does not exist or
        /// is empty, the method returns the string "Null".</remarks>
        /// <param name="path">The path of the registry key to open. This must be a valid registry key path.</param>
        /// <param name="value">The name of the value to retrieve from the specified registry key.</param>
        /// <returns>The string representation of the registry value if it exists and is not empty; otherwise, the string "Null".</returns>
        private static string GetRegValue(string path, string value)
        {
            using RegistryKey? regKey = Registry.LocalMachine.OpenSubKey(path);
            string? str = regKey?.GetValue(value)?.ToString();
            if (string.IsNullOrEmpty(str))
                return "Null";
            else
                return str;
        }

        /// <summary>
        /// Performs a series of tests to verify Office file creation, reading, writing, saving, and deletion rights.
        /// </summary>
        /// <remarks>This method tests the ability to create, read, write, save, and delete a temporary
        /// Word document using Microsoft Office interop. The results of each test are logged to the provided <see
        /// cref="RichTextBox"/> control, and relevant status variables are updated.</remarks>
        /// <param name="rtb">The <see cref="RichTextBox"/> control used to log the results of the tests.</param>
        internal static void OfficeWRTesting(RichTextBox rtb)
        {
            Variables.General_TotalTests += 5;
            Variables.OfficeRights_Hour = DateTime.Now.ToString("HH:mm:ss");

            try
            {
                stopwatch.Restart();

                string fileName = $"temp_{Guid.NewGuid()}.doc";   // Guid named file to avoid collision.
                string filePath = Path.Combine(Path.GetTempPath(), fileName);
                Word.Application wordApp = new()
                {
                    Visible = false
                };

                Word.Document doc = wordApp.Documents.Add();
                doc.Content.Text = "The quick brown fox jumps over the lazy dog";
                doc.SaveAs2(filePath);
                doc.Close();
                if (File.Exists(filePath))
                {
                    rtb.AppendText($"- {T("Create")} : OK" + Environment.NewLine);
                    Variables.OfficeRights_Create = "True";
                    Variables.General_TotalSuccess++;
                }
                else
                {
                    rtb.AppendText($"- {T("Create")} : FAIL." + Environment.NewLine);
                    Variables.OfficeRights_Create = "False";
                    return;
                }

                try
                {
                    doc = wordApp.Documents.Open(filePath);
                    doc.Content.Text += "\nAdding more fox over the lazy dog.";
                    doc.Save();
                    doc.Close();
                }
                catch (Exception ex)
                {
                    rtb.AppendText($"- {T("Write")} / {T("Read")} : FAIL. {ex.Message}" + Environment.NewLine);
                    Variables.OfficeRights_Write = "False";
                    Variables.OfficeRights_Read = "False";
                    Variables.OfficeRights_Save = "False";
                    return;
                }

                doc = wordApp.Documents.Open(filePath);
                if (doc.Content.Text.Contains("Adding more fox over the lazy dog"))
                {
                    rtb.AppendText($"- {T("Save")} : OK" + Environment.NewLine);
                    rtb.AppendText($"- {T("Read")} : OK" + Environment.NewLine);
                    rtb.AppendText($"- {T("Write")} : OK" + Environment.NewLine);
                    Variables.OfficeRights_Save = "True";
                    Variables.OfficeRights_Read = "True";
                    Variables.OfficeRights_Write = "True";
                    Variables.General_TotalSuccess += 3;
                }
                else
                {
                    rtb.AppendText($"- {T("Save")} : FAIL" + Environment.NewLine);
                    rtb.AppendText($"- {T("Read")} : FAIL" + Environment.NewLine);
                    rtb.AppendText($"- {T("Write")} : FAIL" + Environment.NewLine);
                    Variables.OfficeRights_Save = "False";
                    Variables.OfficeRights_Read = "False";
                    Variables.OfficeRights_Write = "False";
                }
                doc.Close();

                wordApp.Quit();
                Marshal.ReleaseComObject(doc);
                Marshal.ReleaseComObject(wordApp);

                File.Delete(filePath);
                if (!File.Exists(filePath))
                {
                    rtb.AppendText($"- {T("Delete")} : OK" + Environment.NewLine);
                    Variables.OfficeRights_Delete = "True";
                    Variables.General_TotalSuccess++;
                }
                else
                {
                    rtb.AppendText($"- {T("Delete")} : FAIL" + Environment.NewLine);
                    Variables.OfficeRights_Delete = "False";
                }
                stopwatch.Stop();
                Variables.OfficeRights_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error OfficeRights: " + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// Tests the availability and status of installed printers on the system and logs the results.
        /// </summary>
        /// <remarks>This method checks for installed printers and retrieves detailed information about
        /// each printer,  including its name, driver, port, and IP address (if available). It also attempts to ping the
        /// printer's IP  to determine its connectivity status. Results are logged to the provided <see
        /// cref="RichTextBox"/> control.  Printers with names containing "Microsoft Print to PDF", "XPS", or "OneNote"
        /// are excluded from the test. If no printers are installed, a message indicating this is logged.</remarks>
        /// <param name="rtb">The <see cref="RichTextBox"/> control where the test results are appended.</param>
        internal static void PrinterTesting(RichTextBox rtb)
        {
            Variables.Printer_Hour = DateTime.Now.ToString("HH:mm:ss");

            try
            {
                stopwatch.Restart();

                if (PrinterSettings.InstalledPrinters.Count == 0)
                {
                    Variables.General_TotalTests++;
                    rtb.AppendText(T("NoPrinterFound") + Environment.NewLine);
                    Variables.General_TotalSuccess++;
                    stopwatch.Stop();
                    Variables.Printer_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
                }
                else
                {
                    // Add printer from installed printers list then do the foreach loop to test each printer.
                    string[] printerCollection = PrinterSettings.InstalledPrinters.Cast<string>().ToArray();
                    printerCollection = printerCollection.Concat(Variables.PrinterList.Split(';').Select(p => p.Trim()).Where(p => !string.IsNullOrEmpty(p))).ToArray();

                    foreach (string printer in printerCollection)
                    {
                        if (!string.IsNullOrWhiteSpace(printer)
                            && !printer.Contains("Fax", StringComparison.OrdinalIgnoreCase)
                            && !printer.Contains("PDF", StringComparison.OrdinalIgnoreCase)
                            && !printer.Contains("Microsoft Print to PDF", StringComparison.OrdinalIgnoreCase)
                            && !printer.Contains("OneNote", StringComparison.OrdinalIgnoreCase)
                            && !printer.Contains("XPS", StringComparison.OrdinalIgnoreCase))
                        {
                            Variables.General_TotalTests++;
                            rtb.AppendText(printer + Environment.NewLine);

                            string printer_clean = printer.Split(',')[0].Trim();
                            printer_clean = printer_clean.Split('(')[0].Trim();
                            printer_clean = printer_clean.Split('\\').Last().Trim();
                            printer_clean = printer_clean.Split('/').Last().Trim();
                            printer_clean = printer_clean.Split(' ').Last().Trim();
                            printer_clean = printer_clean.Trim();

                            Variables.Printer_PrinterName = [.. Variables.Printer_PrinterName, printer_clean];

                            if (IsPrinterReachable_TCP(printer_clean, 9100, Variables.Timeout) || IsPrinterReachable_PING(printer_clean, Variables.Timeout))
                            {
                                Variables.General_TotalSuccess++;
                                Variables.Printer_PrinterStatus = [.. Variables.Printer_PrinterStatus, T("Reachable")];
                                rtb.AppendText($"- Status : {T("Reachable")}" + Environment.NewLine);
                            }
                            else if (IsPrinterConfigurationOK(printer_clean))
                            {
                                Variables.General_TotalSuccess++;
                                Variables.Printer_PrinterStatus = [.. Variables.Printer_PrinterStatus, "Configuration OK"];
                                rtb.AppendText($"- Status : Configuration OK" + Environment.NewLine);
                            }
                            else
                            {
                                Variables.Printer_PrinterStatus = [.. Variables.Printer_PrinterStatus, T("Unreachable")];
                                rtb.AppendText($"- Status : {T("Tests_PrinterError_Configuration")}" + Environment.NewLine);
                            }
                        }
                    }
                }
                stopwatch.Stop();
                Variables.Printer_ElapsedTime = stopwatch.ElapsedMilliseconds.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, T("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Checks if a printer is reachable via TCP connection on the specified port.
        /// </summary>
        /// <param name="ip">The IP address of the printer.</param>
        /// <param name="port">The port number to connect to (default is 9100).</param>
        /// <param name="timeout">The timeout duration in milliseconds (default is 1000).</param>
        /// <returns>True if the printer is reachable; otherwise, false.</returns>
        internal static bool IsPrinterReachable_TCP(string ip, int port = 9100, int timeout = 1000)
        {
            try
            {
                using TcpClient client = new();

                Task connection = client.ConnectAsync(ip, port);

                return Task.WhenAny(
                    connection,
                    Task.Delay(timeout)
                ).ContinueWith(t => t.Result == connection && client.Connected).Result;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Checks if a printer is reachable via ICMP ping.
        /// </summary>
        /// <param name="printerName">The name or IP address of the printer.</param>
        /// <param name="timeout">The timeout duration in milliseconds (default is 1000).</param>
        /// <returns>True if the printer is reachable; otherwise, false.</returns>
        internal static bool IsPrinterReachable_PING(string printerName, int timeout = 1000)
        {
            try
            {
                Ping newPing = new();
                PingReply reply = newPing.Send(printerName, timeout);
                if (reply.Status == IPStatus.Success)
                    return true;
                else
                    return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        internal static bool IsPrinterConfigurationOK(string printerName)
        {
            try
            {
                using ManagementObjectSearcher searcher = new(
                    "SELECT * FROM Win32_Printer WHERE Name = '" +
                    printerName.Replace("'", "''") + "'");

                using ManagementObjectCollection printers = searcher.Get();

                foreach (ManagementObject printer in printers.Cast<ManagementObject>())
                {
                    int score = 0;
                    int maxScore = 0;

                    // Informations essentielles
                    string name = printer["Name"]?.ToString() ?? "";

                    var Printer_ServerName = printer["ServerName"];
                    var Printer_Shared = printer["SharedName"];
                    var Printer_Driver = printer["DriverName"];
                    var Printer_Port = printer["PortName"];
                    var Printer_Location = printer["Location"];

                    // Nom de l'imprimante : 10 points
                    maxScore += 10;
                    if (!string.IsNullOrWhiteSpace(name))
                        score += 10;

                    // Print Server : 25 points
                    maxScore += 25;
                    if (!string.IsNullOrWhiteSpace(Printer_ServerName?.ToString()))
                        score += 25;

                    // ShareName : 20 points
                    maxScore += 20;
                    if (!string.IsNullOrWhiteSpace(Printer_Shared?.ToString()))
                        score += 20;

                    // Driver : 15 points
                    maxScore += 15;
                    if (!string.IsNullOrWhiteSpace(Printer_Driver?.ToString()))
                        score += 15;

                    // Port : 15 points
                    maxScore += 15;
                    if (!string.IsNullOrWhiteSpace(Printer_Port?.ToString()))
                        score += 15;

                    // Type d'imprimante
                    bool network = printer["Network"] as bool? ?? false;
                    bool local = printer["Local"] as bool? ?? false;
                    bool shared = printer["Shared"] as bool? ?? false;

                    // Imprimante réseau : 5 points
                    maxScore += 5;
                    if (network)
                        score += 5;

                    // Imprimante partagée : 5 points
                    maxScore += 5;
                    if (shared)
                        score += 5;

                    // Pas une imprimante locale : 5 points
                    maxScore += 5;
                    if (!local)
                        score += 5;

                    // Location : 5 points
                    maxScore += 5;
                    if (!string.IsNullOrWhiteSpace(Printer_Location?.ToString()))
                        score += 5;

                    // PrinterStatus :
                    // 1 = Other
                    // 2 = Unknown
                    // 3 = Idle
                    // 4 = Printing
                    // 5 = Warming Up
                    // 6 = Stopped Printing
                    // 7 = Offline
                    int printerStatus = Convert.ToInt32(
                        printer["PrinterStatus"] ?? 0);

                    maxScore += 5;

                    switch (printerStatus)
                    {
                        case 3: // Idle
                        case 4: // Printing
                            score += 5;
                            break;

                        case 5: // Warming Up
                            score += 3;
                            break;

                        case 1: // Other
                        case 6: // Stopped Printing
                            score += 1;
                            break;

                        case 2: // Unknown
                        case 7: // Offline
                        default:
                            break;
                    }

                    double percentage = (double)score / maxScore * 100.0;

                    return percentage >= 80.0;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Checks if the provided URL is valid and uses either the HTTP or HTTPS scheme.
        /// </summary>
        /// <param name="url">The URL to check.</param>
        /// <returns>True if the URL is valid and uses HTTP or HTTPS; otherwise, false.</returns>
        internal static bool Check_URL(string url)
        {
            if (Uri.TryCreate(url, UriKind.Absolute, out Uri? uriResult) &&
                (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps))
                return true;
            else
                return false;
        }
    }
}
