namespace AccountTester
{
    internal class ExportVariables
    {
        // Variables for ExportForm, valus are set by the MainForm functions

        // General variables
        public static string General_export_DeviceOS = Convert.ToUInt32(Environment.OSVersion.Version.ToString().Split('.')[2]) >= 22631 ? "Windows 11" : "Windows 10";

        public static string General_export_OSArchitecture = Environment.Is64BitOperatingSystem ? "64 bits" : "32 bits" ?? "Error";

        public static string General_export_UserName = Environment.UserName;

        public static string General_DateAndHour = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        public static int General_export_TotalTests { get; set; }
        public static int General_export_TotalSuccess { get; set; }
        public static string General_export_Resume { get; set; } = "Null";

        // Internet Connection variables
        public static string InternetConnexion_export_Hour { get; set; } = "Null";

        public static string InternetConnexion_export_TestedURL = "https://www.google.ch";
        public static string InternetConnexion_export_HTMLStatut { get; set; } = "Null";
        public static string InternetConnexion_export_ElapsedTime { get; set; } = "Null";

        // NetworkStorageRights variables - Not OK, need to be tested with a network share
        public static string NetworkStorageRights_export_Hour { get; set; } = "Null";
        public static string NetworkStorageRights_export_ConnexionType { get; set; } = "Null"; // Not set
        public static string[]? NetworkStorageRights_export_DiskLetter { get; set; }  // Not set
        public static string[]? NetworkStorageRights_export_CheminUNC { get; set; } // Not set
        public static string[]? NetworkStorageRights_export_Serveur { get; set; }   // Not set
        public static string[]? NetworkStorageRights_export_ShareName { get; set; } // Not set
        public static string NetworkStorageRights_export_ElapsedTime { get; set; } = "Null";

        // OfficeVersion variables
        public static string OfficeVersion_export_Hour { get; set; } = "Null";
        public static string OfficeVersion_export_OfficeVersion { get; set; } = "Null";
        public static string OfficeVersion_export_OfficePath { get; set; } = "Null";
        public static string OfficeVersion_export_OfficeCulture { get; set; } = "Null";
        public static string OfficeVersion_export_OfficeExcludedApps { get; set; } = "Null";
        public static string OfficeVersion_export_OfficeLastUpdateStatus { get; set; } = "Null";
        public static string OfficeVersion_export_ElapsedTime { get; set; } = "Null";

        // OfficeRights variables
        public static string OfficeRights_export_Hour { get; set; } = "Null";
        public static string OfficeRights_export_Write { get; set; } = "Null";
        public static string OfficeRights_export_Read { get; set; } = "Null";
        public static string OfficeRights_export_Delete { get; set; } = "Null";
        public static string OfficeRights_export_Save { get; set; } = "Null";
        public static string OfficeRights_export_Create { get; set; } = "Null";

        public static string OfficeRights_export_FolderTested = Path.GetTempPath();
        public static string OfficeRights_export_ElapsedTime { get; set; } = "Null";

        // Printer variables
        public static string Printer_export_Hour { get; set; } = "Null";
        public static string[] Printer_export_PrinterName { get; set; } = [];
        public static string[] Printer_export_PrinterIP { get; set; } = [];
        public static string[] Printer_export_PrinterStatus { get; set; } = [];
        public static string[] Printer_export_PrinterDriver { get; set; } = [];
        public static string[] Printer_export_PrinterPort { get; set; } = [];
        public static string Printer_export_ElapsedTime { get; set; } = "Null";
    }
}
