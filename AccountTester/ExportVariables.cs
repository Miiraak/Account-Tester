namespace AccountTester
{
    internal class ExportVariables
    {
        // Variables for ExportForm, valus are set by the MainForm.cs

        // General variables - OK
        public static string General_export_DeviceOS = Convert.ToUInt32(Environment.OSVersion.Version.ToString().Split('.')[2]) >= 22631 ? "Windows 11" : "Windows 10";   // OK
        public static string? General_export_ProcessType = Environment.Is64BitProcess ? "64 bits" : "32 bits";     // OK
        public static string? General_export_OSArchitecture = Environment.Is64BitOperatingSystem ? "64 bits" : "32 bits";     // OK
        public static string General_export_UserName = Environment.UserName;     // OK

        public static int General_export_TotalTests { get; set; } = 0;
        public static int General_export_TotalSuccess { get; set; } = 0;

        public static string? General_export_Resume { get; set; }  // OK

        // Internet Connection variables - Error
        public static string? InternetConnexion_export_DateAndHour { get; set; }    // OK
        public static string InternetConnexion_export_TestedURL = "https://www.google.ch";   // OK
        public static string? InternetConnexion_export_HTMLStatut { get; set; }      // OK
        public static string? InternetConnexion_export_ElapsedTime { get; set; }    // OK

        // NetworkStorageRights variables 
        public static string? NetworkStorageRights_export_DateAndHour { get; set; }     // OK
        public static string? NetworkStorageRights_export_ConnexionType { get; set; }
        public static string[]? NetworkStorageRights_export_DiskLetter { get; set; }
        public static string[]? NetworkStorageRights_export_CheminUNC { get; set; }
        public static string[]? NetworkStorageRights_export_Serveur { get; set; }
        public static string[]? NetworkStorageRights_export_ShareName { get; set; }
        public static string? NetworkStorageRights_export_ElapsedTime { get; set; } // Incorrect si aucun disque réseau

        // OfficeVersion variables
        public static string? OfficeVersion_export_DateAndHour { get; set; }
        public static string? OfficeVersion_export_OfficeVersion { get; set; }
        public static string? OfficeVersion_export_OfficeEdition { get; set; }
        public static string? OfficeVersion_export_OfficeArchitecture { get; set; }
        public static string? OfficeVersion_export_OfficePath { get; set; }
        public static string? OfficeVersion_export_OfficeProductID { get; set; }
        public static string? OfficeVersion_export_OfficeSerialNumber { get; set; }
        public static string? OfficeVersion_export_OfficeSerialNumberStatus { get; set; }
        public static string? OfficeVersion_export_ElapsedTime { get; set; }    // Incorrect n'affiche pas le temps

        // OfficeRights variables
        public static string? OfficeRights_export_DateAndHour { get; set; }      // set
        public static string? OfficeRights_export_CanWrite { get; set; }    // Set
        public static string? OfficeRights_export_CanRead { get; set; }      // Set
        public static string? OfficeRights_export_CanDelete { get; set; }      // set
        public static string? OfficeRights_export_CanCopy { get; set; }
        public static string? OfficeRights_export_CanMove { get; set; }
        public static string? OfficeRights_export_CanRename { get; set; }
        public static string? OfficeRights_export_CanSave { get; set; }            // set
        public static string? OfficeRights_export_CanCreate { get; set; }     // Set
        public static string[]? OfficeRights_export_FolderTested { get; set; }  // Fonctionne si aucun disque, a tester avec un disque réseau
        public static string? OfficeRights_export_ElapsedTime { get; set; }     // OK

        // Printer variables
        public static string? Printer_export_DateAndHour { get; set; }
        public static string[]? Printer_export_PrinterName { get; set; }      // Incorrect
        public static string[]? Printer_export_PrinterStatus { get; set; }
        public static string[]? Printer_export_PrinterDriver { get; set; }
        public static string[]? Printer_export_PrinterPort { get; set; }
        public static string? Printer_export_ElapsedTime { get; set; }      // OK
    }
}
