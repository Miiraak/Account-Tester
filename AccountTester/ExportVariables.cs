namespace AccountTester
{
    internal class ExportVariables
    {
        // Variables for ExportForm, valus are set by the MainForm functions

        // General variables
        public static string General_DeviceOS = Convert.ToUInt32(Environment.OSVersion.Version.ToString().Split('.')[2]) >= 22631 ? "Windows 11" : "Windows 10";

        public static string General_OSArchitecture = Environment.Is64BitOperatingSystem ? "64 bits" : "32 bits" ?? "Error";

        public static string General_UserName = Environment.UserName;

        public static string General_DateAndHour = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        public static int General_TotalTests { get; set; }
        public static int General_TotalSuccess { get; set; }
        public static string General_Resume { get; set; } = "Null";

        // Internet Connection variables
        public static string InternetConnexion_Hour { get; set; } = "Null";

        public static string InternetConnexion_TestedURL = "https://www.google.ch";
        public static string InternetConnexion_HTMLStatut { get; set; } = "Null";
        public static string InternetConnexion_ElapsedTime { get; set; } = "Null";

        // NetworkStorageRights variables - Not OK and need to be tested with a network share 
        public static string NetworkStorageRights_Hour { get; set; } = "Null";
        public static string[] NetworkStorageRights_DiskLetter { get; set; } = [];  // Ok with local disk, not test with network share
        public static string[] NetworkStorageRights_CheminUNC { get; set; } = Array.Empty<string>(); // OK with local disk, not test with network share
        public static string[] NetworkStorageRights_Serveur { get; set; } = [];   // Ok with local disk, not test with network share
        public static string[] NetworkStorageRights_ShareName { get; set; } = []; // Ok with local disk, not test with network share
        public static string NetworkStorageRights_ElapsedTime { get; set; } = "Null";

        // OfficeVersion variables
        public static string OfficeVersion_Hour { get; set; } = "Null";
        public static string OfficeVersion_OfficeVersion { get; set; } = "Null";
        public static string OfficeVersion_OfficePath { get; set; } = "Null";
        public static string OfficeVersion_OfficeCulture { get; set; } = "Null";
        public static string OfficeVersion_OfficeExcludedApps { get; set; } = "Null";
        public static string OfficeVersion_OfficeLastUpdateStatus { get; set; } = "Null";
        public static string OfficeVersion_ElapsedTime { get; set; } = "Null";

        // OfficeRights variables
        public static string OfficeRights_Hour { get; set; } = "Null";
        public static string OfficeRights_Write { get; set; } = "Null";
        public static string OfficeRights_Read { get; set; } = "Null";
        public static string OfficeRights_Delete { get; set; } = "Null";
        public static string OfficeRights_Save { get; set; } = "Null";
        public static string OfficeRights_Create { get; set; } = "Null";

        public static string OfficeRights_FolderTested = Path.GetTempPath();
        public static string OfficeRights_ElapsedTime { get; set; } = "Null";

        // Printer variables
        public static string Printer_Hour { get; set; } = "Null";
        public static string[] Printer_PrinterName { get; set; } = [];
        public static string[] Printer_PrinterIP { get; set; } = [];
        public static string[] Printer_PrinterStatus { get; set; } = [];
        public static string[] Printer_PrinterDriver { get; set; } = [];
        public static string[] Printer_PrinterPort { get; set; } = [];
        public static string Printer_ElapsedTime { get; set; } = "Null";
    }
}
