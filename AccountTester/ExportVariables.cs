using System.Net.NetworkInformation;

namespace AccountTester
{
    internal class ExportVariables
    {
        // Variables for ExportForm, valus are set by the MainForm.cs

        // General variables
        public static string General_export_DeviceOS = Environment.OSVersion.ToString();      // Incorrect
        public static string? General_export_ProcessType = Environment.Is64BitProcess ? "64 bits" : "32 bits";     // OK
        public static string? General_export_OSArchitecture = Environment.Is64BitOperatingSystem ? "64 bits" : "32 bits";     // OK
        public static string General_export_UserName = Environment.UserName;     // OK
        public static string? General_export_Resume { get; set; }  // OK

        // Internet Connection variables
        public static string? InternetConnexion_export_DateAndHour { get; set; }    // OK
        public static string[]? InternetConnexion_export_ConnexionType = SetConnexionType();  // To Test
        public static string InternetConnexion_export_TestedURL = "https://www.google.ch";   // OK
        public static string? InternetConnexion_export_HTMLStatut { get; set; }     // OK
        public static string? InternetConnexion_export_ResponseTime { get; set; }

        // NetworkStorageRights variables 
        public static string? NetworkStorageRights_export_DateAndHour { get; set; }     // OK
        public static string? NetworkStorageRights_export_UsedProtocol { get; set; }
        public static string? NetworkStorageRights_export_ConnexionType { get; set; }
        public static string[]? NetworkStorageRights_export_DiskLetter { get; set; }
        public static string[]? NetworkStorageRights_export_CheminUNC { get; set; }
        public static string[]? NetworkStorageRights_export_Serveur { get; set; }
        public static string[]? NetworkStorageRights_export_ShareName { get; set; }

        // OfficeVersion variables
        public static string? OfficeVersion_export_DateAndHour { get; set; }
        public static string? OfficeVersion_export_OfficeVersion { get; set; }
        public static string? OfficeVersion_export_OfficeEdition { get; set; }
        public static string? OfficeVersion_export_OfficeArchitecture { get; set; }
        public static string? OfficeVersion_export_OfficePath { get; set; }
        public static string? OfficeVersion_export_OfficeProductID { get; set; }
        public static string? OfficeVersion_export_OfficeSerialNumber { get; set; }
        public static string? OfficeVersion_export_OfficeSerialNumberStatus { get; set; }

        // OfficeRights variables
        public static string? OfficeRights_export_DateAndHour { get; set; }
        public static string? OfficeRights_export_CanWrite { get; set; }
        public static string? OfficeRights_export_CanRead { get; set; }
        public static string? OfficeRights_export_CanExecute { get; set; }
        public static string? OfficeRights_export_CanDelete { get; set; }
        public static string? OfficeRights_export_CanCopy { get; set; }
        public static string? OfficeRights_export_CanMove { get; set; }
        public static string? OfficeRights_export_CanRename { get; set; }
        public static string? OfficeRights_export_CanCreate { get; set; }
        public static string[]? OfficeRights_export_FolderTested { get; set; }

        // Printer variables
        public static string? Printer_export_DateAndHour { get; set; }
        public static string[]? Printer_export_PrinterName { get; set; }
        public static string[]? Printer_export_PrinterStatus { get; set; }
        public static string[]? Printer_export_PrinterDriver { get; set; }
        public static string[]? Printer_export_PrinterPort { get; set; }
        public static string[]? Printer_export_PrinterLocation { get; set; }

        /// <summary>
        /// Get the connexion type of the network interface
        /// </summary>
        /// <returns>A type list of all NetworkInterface where status is Up</returns>
        public static string[] SetConnexionType()
        {
            string[] InternetConnexionTypes = new string[NetworkInterface.GetAllNetworkInterfaces().Length];
            int i = 0;
            foreach (NetworkInterface netInterface in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (netInterface.OperationalStatus == OperationalStatus.Up)
                {
                    InternetConnexionTypes[i] = netInterface.NetworkInterfaceType.ToString();
                    i++;
                }
            }
            return InternetConnexionTypes;
        }
    }


}
