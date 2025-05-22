namespace AccountTester
{
    internal class ExportVariables
    {
        // Variables for ExportForm, valus are set by the MainForm.cs

        // General variables
        public static string General_export_DeviceOS = Convert.ToUInt32(Environment.OSVersion.Version.ToString().Split('.')[2]) >= 22631 ? "Windows 11" : "Windows 10";  // OK 

        public static string General_export_OSArchitecture = Environment.Is64BitOperatingSystem ? "64 bits" : "32 bits" ?? "Error";  // OK

        public static string General_export_UserName = Environment.UserName; // OK

        public static string General_DateAndHour = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"); // OK
        public static int General_export_TotalTests { get; set; } // Not right value in report
        public static int General_export_TotalSuccess { get; set; } // Not right value in report
        public static string General_export_Resume { get; set; } = "Null"; // OK

        // Internet Connection variables
        public static string InternetConnexion_export_Hour { get; set; } = "Null";    // OK

        public static string InternetConnexion_export_TestedURL = "https://www.google.ch"; // OK
        public static string InternetConnexion_export_HTMLStatut { get; set; } = "Null";  // OK
        public static string InternetConnexion_export_ElapsedTime { get; set; } = "Null"; // OK

        // NetworkStorageRights variables 
        public static string NetworkStorageRights_export_Hour { get; set; } = "Null";     // OK
        public static string NetworkStorageRights_export_ConnexionType { get; set; } = "Null"; // Not set
        public static string[]? NetworkStorageRights_export_DiskLetter { get; set; }  // Not set
        public static string[]? NetworkStorageRights_export_CheminUNC { get; set; } // Not set
        public static string[]? NetworkStorageRights_export_Serveur { get; set; } // Not set
        public static string[]? NetworkStorageRights_export_ShareName { get; set; } // Not set
        public static string NetworkStorageRights_export_ElapsedTime { get; set; } = "Null"; // OK

        // OfficeVersion variables
        public static string OfficeVersion_export_Hour { get; set; } = "Null"; // OK
        public static string OfficeVersion_export_OfficeVersion { get; set; } = "Null";  // OK
        public static string OfficeVersion_export_OfficePath { get; set; } = "Null"; // Not set
        public static string OfficeVersion_export_ElapsedTime { get; set; } = "Null";    // OK

        // OfficeRights variables - OK
        public static string OfficeRights_export_Hour { get; set; } = "Null";      // OK
        public static string OfficeRights_export_CanWrite { get; set; } = "Null";    // OK
        public static string OfficeRights_export_CanRead { get; set; } = "Null";      // OK
        public static string OfficeRights_export_CanDelete { get; set; } = "Null";      // OK
        public static string OfficeRights_export_CanSave { get; set; } = "Null";     // OK
        public static string OfficeRights_export_CanCreate { get; set; } = "Null";     // OK

        public static string OfficeRights_export_FolderTested = Path.GetTempPath();  // OK
        public static string OfficeRights_export_ElapsedTime { get; set; } = "Null";     // OK

        // Printer variables
        public static string Printer_export_Hour { get; set; } = "Null";     // Ok
        public static string[]? Printer_export_PrinterName { get; set; }      // Incorrect in report
        public static string[]? Printer_export_PrinterStatus { get; set; }
        public static string[]? Printer_export_PrinterDriver { get; set; }
        public static string[]? Printer_export_PrinterPort { get; set; }
        public static string Printer_export_ElapsedTime { get; set; } = "Null";      // OK
    }
}
