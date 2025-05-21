namespace AccountTester
{
    internal class ExportVariables
    {
        // Variables for ExportForm, valus are set by the MainForm.cs

        // General variables - OK
        public static string General_export_DeviceOS = Convert.ToUInt32(Environment.OSVersion.Version.ToString().Split('.')[2]) >= 22631 ? "Windows 11" : "Windows 10";   // OK
        public static string General_export_OSArchitecture = Environment.Is64BitOperatingSystem ? "64 bits" : "32 bits" ?? "Error";     // OK
        public static string General_export_UserName = Environment.UserName;     // OK
        public static string General_DateAndHour = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");   // OK

        public static int General_export_TotalTests { get; set; } // OK
        public static int General_export_TotalSuccess { get; set; } // OK
        public static string General_export_Resume { get; set; } = "Error"; // OK

        // Internet Connection variables - OK
        public static string InternetConnexion_export_Hour { get; set; } = "Error";    // OK
        public static string InternetConnexion_export_TestedURL = "https://www.google.ch";   // OK
        public static string InternetConnexion_export_HTMLStatut { get; set; } = "Error";      // OK
        public static string InternetConnexion_export_ElapsedTime { get; set; } = "Error";    // OK

        // NetworkStorageRights variables 
        public static string NetworkStorageRights_export_Hour { get; set; } = "Error";     // OK
        public static string NetworkStorageRights_export_ConnexionType { get; set; } = "Error";
        public static string[]? NetworkStorageRights_export_DiskLetter { get; set; }
        public static string[]? NetworkStorageRights_export_CheminUNC { get; set; }
        public static string[]? NetworkStorageRights_export_Serveur { get; set; }
        public static string[]? NetworkStorageRights_export_ShareName { get; set; }
        public static string NetworkStorageRights_export_ElapsedTime { get; set; } = "Error"; // OK

        // OfficeVersion variables
        public static string OfficeVersion_export_Hour { get; set; } = "Error";
        public static string OfficeVersion_export_OfficeVersion { get; set; } = "Error";
        public static string OfficeVersion_export_OfficePath { get; set; } = "Error";
        public static string OfficeVersion_export_ElapsedTime { get; set; } = "Error";    // OK

        // OfficeRights variables
        public static string OfficeRights_export_Hour { get; set; } = "Error";      // OK
        public static string OfficeRights_export_CanWrite { get; set; } = "Error";    // OK
        public static string OfficeRights_export_CanRead { get; set; } = "Error";      // OK
        public static string OfficeRights_export_CanDelete { get; set; } = "Error";      // OK
        public static string OfficeRights_export_CanSave { get; set; } = "Error";     // OK
        public static string OfficeRights_export_CanCreate { get; set; } = "Error";     // OK
        public static string OfficeRights_export_FolderTested = Path.GetTempPath();  // Fonctionne si aucun disque, a tester avec un disque réseau
        public static string OfficeRights_export_ElapsedTime { get; set; } = "Error";     // OK

        // Printer variables
        public static string Printer_export_Hour { get; set; } = "Error";     // Ok
        public static string[]? Printer_export_PrinterName { get; set; }      // Incorrect
        public static string[]? Printer_export_PrinterStatus { get; set; }
        public static string[]? Printer_export_PrinterDriver { get; set; }
        public static string[]? Printer_export_PrinterPort { get; set; }
        public static string Printer_export_ElapsedTime { get; set; } = "Error";      // OK
    }
}
