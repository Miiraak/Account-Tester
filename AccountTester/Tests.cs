using Microsoft.Win32;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace AccountTester
{
    internal class Tests
    {
        static string T(string key) => LangManager.Instance.Translate(key);

        static Stopwatch stopwatch = new();

        /// <summary>
        /// Tests the internet connection by sending an HTTP GET request to a predefined URL.
        /// </summary>
        /// <remarks>This method performs an asynchronous HTTP GET request to the URL specified in
        /// <c>Variables.InternetConnexion_TestedURL</c>. It logs the connection status and response details to the
        /// provided <see cref="RichTextBox"/>. The method updates several global variables, including the total number
        /// of tests, the elapsed time for the test, and the HTTP status code of the response.</remarks>
        /// <param name="rtb">The <see cref="RichTextBox"/> used to display logs related to the connection test.</param>
        /// <returns></returns>
        internal static async Task InternetConnexionTest(RichTextBox rtb)
        {
            Variables.General_TotalTests++;

            try
            {
                stopwatch.Restart();
                string Target = Variables.Target;
                using HttpClient client = new();
                client.Timeout = TimeSpan.FromSeconds(Variables.Timeout);

                // Check if the target URL starts with "http://" or "https://" if not, prepend "http://" to it.
                if (!Target.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                {
                    Target = "http://" + Variables.Target;
                }

                HttpResponseMessage response = await client.GetAsync(Target);

                Variables.InternetConnexion_Hour = DateTime.Now.ToString("HH:mm:ss");
                Variables.InternetConnexion_HTMLStatut = response.StatusCode.ToString();

                if (response.IsSuccessStatusCode)
                {
                    rtb.AppendText($"{T("MainForm_RTBL_Internet_Connected")}" + Environment.NewLine);
                    Variables.General_TotalSuccess++;
                }
                else
                {
                    rtb.AppendText($"{T("MainForm_RTBL_Internet_Others")}" + response.StatusCode + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                rtb.AppendText($"{T("MainForm_RTBL_Internet_Others")}" + ex.InnerException?.Message + Environment.NewLine);
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
                    foundDrives = foundDrives.Append(drive.Name[0].ToString()).ToArray();
                    Variables.General_TotalTests++;
                    Variables.NetworkStorageRights_DiskLetter = Variables.NetworkStorageRights_DiskLetter.Append(drive.Name).ToArray();

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
                                Variables.NetworkStorageRights_CheminUNC = Variables.NetworkStorageRights_CheminUNC.Append(cheminUNC).ToArray();
                                Variables.NetworkStorageRights_Serveur = Variables.NetworkStorageRights_Serveur.Append(serveur).ToArray();
                                Variables.NetworkStorageRights_ShareName = Variables.NetworkStorageRights_ShareName.Append(shareName).ToArray();
                            }

                            File.Delete(testFile);
                        }
                        catch (UnauthorizedAccessException)
                        {
                            rtb.AppendText($@"- {drive.Name} : {T("MainForm_RTBL_NetworkStorageRightsTesting_Refused")}" + Environment.NewLine);
                            Variables.NetworkStorageRights_CheminUNC = Variables.NetworkStorageRights_CheminUNC.Append(T("UnauthorizedAccess")).ToArray();
                            Variables.NetworkStorageRights_Serveur = Variables.NetworkStorageRights_Serveur.Append(T("UnauthorizedAccess")).ToArray();
                            Variables.NetworkStorageRights_ShareName = Variables.NetworkStorageRights_ShareName.Append(T("UnauthorizedAccess")).ToArray();
                        }
                        catch (IOException)
                        {
                            rtb.AppendText($@"- {drive.Name} : {T("MainForm_RTBL_NetworkStorageRightsTesting_Error")}" + Environment.NewLine);
                            Variables.NetworkStorageRights_CheminUNC = Variables.NetworkStorageRights_CheminUNC.Append(T("IOError")).ToArray();
                            Variables.NetworkStorageRights_Serveur = Variables.NetworkStorageRights_Serveur.Append(T("IOError")).ToArray();
                            Variables.NetworkStorageRights_ShareName = Variables.NetworkStorageRights_ShareName.Append(T("IOError")).ToArray();
                        }
                    }
                    else
                    {
                        Variables.NetworkStorageRights_CheminUNC = Variables.NetworkStorageRights_CheminUNC.Append(drive.Name).ToArray();
                        Variables.NetworkStorageRights_Serveur = Variables.NetworkStorageRights_Serveur.Append("localhost").ToArray();
                        Variables.NetworkStorageRights_ShareName = Variables.NetworkStorageRights_ShareName.Append(T("None")).ToArray();

                        rtb.AppendText($@"- {drive.Name} : {T("Omitted")}" + Environment.NewLine);
                        Variables.General_TotalSuccess++;
                    }
                }

                string[] drivesList = Variables.DrivesList.Split(';').Select(p => p.Trim()).Where(p => !string.IsNullOrEmpty(p)).ToArray();
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

                doc = wordApp.Documents.Open(filePath);
                doc.Content.Text += "\nAdding more fox over the lazy dog.";
                doc.Save();
                doc.Close();

                doc = wordApp.Documents.Open(filePath);
                if (doc.Content.Text.Contains("Adding more fox over the lazy dog"))
                {
                    rtb.AppendText($"- {T("Save")} : OK" + Environment.NewLine);
                    Variables.OfficeRights_Save = "True";
                    Variables.General_TotalSuccess++;
                }
                else
                {
                    rtb.AppendText($"- {T("Save")} : FAIL" + Environment.NewLine);
                    Variables.OfficeRights_Save = "False";
                }
                doc.Close();

                doc = wordApp.Documents.Open(filePath);
                if (doc.Content.Text.Contains("The quick brown fox jumps over the lazy dog"))
                {
                    rtb.AppendText($"- {T("Read")} : OK" + Environment.NewLine);
                    rtb.AppendText($"- {T("Write")} : OK" + Environment.NewLine);
                    Variables.OfficeRights_Read = "True";
                    Variables.OfficeRights_Write = "True";
                    Variables.General_TotalSuccess += 2;
                }
                else
                {
                    rtb.AppendText($"- {T("Read")} : FAIL" + Environment.NewLine);
                    rtb.AppendText($"- {T("Write")} : FAIL" + Environment.NewLine);
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
                    string[] foundPrinter = [];

                    foreach (string printerName in PrinterSettings.InstalledPrinters)
                    {
                        string printer = printerName;
                        if (printer.Contains('\\', StringComparison.Ordinal))
                            printer = printer.Split('\\').Last();

                        if (!printer.Contains("Microsoft Print to PDF", StringComparison.OrdinalIgnoreCase) &&
                            !printer.Contains("XPS", StringComparison.OrdinalIgnoreCase) &&
                            !printer.Contains("OneNote", StringComparison.OrdinalIgnoreCase))
                        {
                            foundPrinter = foundPrinter.Append(printer).ToArray();
                            Variables.General_TotalTests++;
                            string registryPath = @"SYSTEM\CurrentControlSet\Control\Print\Printers\" + printer;

                            using RegistryKey? printerKey = Registry.LocalMachine.OpenSubKey(registryPath);
                            if (printerKey != null)
                            {
                                Variables.Printer_PrinterName = Variables.Printer_PrinterName.Append(printer).ToArray();
                                Variables.Printer_PrinterDriver = Variables.Printer_PrinterDriver.Append(printerKey.GetValue("Printer Driver")?.ToString() ?? T("Unknown")).ToArray();
                                Variables.Printer_PrinterPort = Variables.Printer_PrinterPort.Append(printerKey.GetValue("Port")?.ToString() ?? T("Unknown")).ToArray();

                                string? locationValue = printerKey.GetValue("Location")?.ToString();
                                if (!string.IsNullOrEmpty(locationValue))
                                {
                                    string PrinterIP = locationValue.Split("//").Last().Split(":").First();
                                    Variables.Printer_PrinterIP = Variables.Printer_PrinterIP.Append(PrinterIP).ToArray();

                                    if (!string.IsNullOrEmpty(PrinterIP))
                                    {
                                        Ping ping = new();
                                        PingReply reply = ping.Send(PrinterIP, 1000);

                                        if (reply.Status == IPStatus.Success)
                                        {
                                            rtb.AppendText(printer + Environment.NewLine);
                                            rtb.AppendText("- IP : " + PrinterIP + Environment.NewLine + "- Ping : OK" + Environment.NewLine);
                                            Variables.Printer_PrinterStatus = Variables.Printer_PrinterStatus.Append("OK").ToArray();
                                            Variables.General_TotalSuccess++;
                                        }
                                        else
                                        {
                                            rtb.AppendText(printer + Environment.NewLine);
                                            rtb.AppendText("- IP : " + PrinterIP + Environment.NewLine + "- Ping : FAIL" + Environment.NewLine);
                                            Variables.Printer_PrinterStatus = Variables.Printer_PrinterStatus.Append("FAIL").ToArray();
                                        }
                                    }
                                    else
                                    {
                                        rtb.AppendText(printer + Environment.NewLine);
                                        rtb.AppendText($"- IP : {T("MainForm_RTBL_PrinterTesting_NotFound")}" + Environment.NewLine);
                                        Variables.Printer_PrinterIP = Variables.Printer_PrinterIP.Append(T("Unknown")).ToArray();
                                        Variables.Printer_PrinterStatus = Variables.Printer_PrinterStatus.Append(T("Unknown")).ToArray();
                                        Variables.Printer_PrinterDriver = Variables.Printer_PrinterDriver.Append(T("Unknown")).ToArray();
                                        Variables.Printer_PrinterPort = Variables.Printer_PrinterPort.Append(T("Unknown")).ToArray();
                                    }
                                }
                                else
                                {
                                    rtb.AppendText(printer + Environment.NewLine);
                                    rtb.AppendText($"- {T("MainForm_RTBL_PrinterTesting_NoLocationValueReg")}" + Environment.NewLine);
                                    Variables.Printer_PrinterIP = Variables.Printer_PrinterIP.Append(T("Unknown")).ToArray();
                                    Variables.Printer_PrinterStatus = Variables.Printer_PrinterStatus.Append(T("Unknown")).ToArray();
                                    Variables.Printer_PrinterDriver = Variables.Printer_PrinterDriver.Append(T("Unknown")).ToArray();
                                    Variables.Printer_PrinterPort = Variables.Printer_PrinterPort.Append(T("Unknown")).ToArray();
                                }
                            }
                            else
                            {
                                rtb.AppendText(printer + Environment.NewLine);
                                rtb.AppendText($"- {T("MainForm_RTBL_NoRegKey")}" + Environment.NewLine);
                                Variables.Printer_PrinterIP = Variables.Printer_PrinterIP.Append(T("Unknown")).ToArray();
                                Variables.Printer_PrinterStatus = Variables.Printer_PrinterStatus.Append(T("Unknown")).ToArray();
                                Variables.Printer_PrinterDriver = Variables.Printer_PrinterDriver.Append(T("Unknown")).ToArray();
                                Variables.Printer_PrinterPort = Variables.Printer_PrinterPort.Append(T("Unknown")).ToArray();
                            }
                        }
                    }

                    string[] printerList = Variables.PrinterList.Split(';').Select(p => p.Trim()).Where(p => !string.IsNullOrEmpty(p)).ToArray();
                    foreach (string printer in printerList)
                    {
                        if (!foundPrinter.Contains(printer))
                        {
                            rtb.AppendText($"{printer} ({T("Missing")})" + Environment.NewLine);
                            rtb.AppendText($"- {T("NoPrinterFound")}" + Environment.NewLine);
                            Variables.General_TotalTests++;
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
    }
}
