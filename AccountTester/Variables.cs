namespace AccountTester
{
    internal class Variables
    {
        // Variables for ExportForm, values are set by the MainForm functions
        // General variables
        public static string General_DeviceOS = Convert.ToUInt32(Environment.OSVersion.Version.ToString().Split('.')[2]) >= 22631 ? "Windows 11" : "Windows 10";

        public static string General_OSArchitecture = Environment.Is64BitOperatingSystem ? "64 bits" : "32 bits" ?? "Error";

        public static string General_UserName = Environment.UserName;

        public static string General_DateAndHour = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
        public static int General_TotalTests { get; set; } = 0;
        public static int General_TotalSuccess { get; set; } = 0;
        public static string General_Resume { get; set; } = string.Empty;

        // Internet Connection variables
        public static string InternetConnexion_Hour { get; set; } = string.Empty;
        public static string InternetConnexion_HTMLStatut { get; set; } = string.Empty;
        public static string InternetConnexion_ElapsedTime { get; set; } = string.Empty;

        // NetworkStorageRights variables
        public static string NetworkStorageRights_Hour { get; set; } = string.Empty;
        public static string[] NetworkStorageRights_DiskLetter { get; set; } = [];
        public static string[] NetworkStorageRights_CheminUNC { get; set; } = [];
        public static string[] NetworkStorageRights_Serveur { get; set; } = [];
        public static string[] NetworkStorageRights_ShareName { get; set; } = [];
        public static string NetworkStorageRights_ElapsedTime { get; set; } = string.Empty;

        // OfficeVersion variables
        public static string OfficeVersion_Hour { get; set; } = string.Empty;
        public static string OfficeVersion_OfficeVersion { get; set; } = string.Empty;
        public static string OfficeVersion_OfficePath { get; set; } = string.Empty;
        public static string OfficeVersion_OfficeCulture { get; set; } = string.Empty;
        public static string OfficeVersion_OfficeExcludedApps { get; set; } = string.Empty;
        public static string OfficeVersion_OfficeLastUpdateStatus { get; set; } = string.Empty;
        public static string OfficeVersion_ElapsedTime { get; set; } = string.Empty;

        // OfficeRights variables
        public static string OfficeRights_Hour { get; set; } = string.Empty;
        public static string OfficeRights_Write { get; set; } = string.Empty;
        public static string OfficeRights_Read { get; set; } = string.Empty;
        public static string OfficeRights_Delete { get; set; } = string.Empty;
        public static string OfficeRights_Save { get; set; } = string.Empty;
        public static string OfficeRights_Create { get; set; } = string.Empty;

        public static string OfficeRights_FolderTested = Path.GetTempPath();
        public static string OfficeRights_ElapsedTime { get; set; } = string.Empty;

        // Printer variables
        public static string[] PrinterSystemList { get; set; } = []; // List of printers that appear in the system, used to check if the printer is installed or not.
        public static string Printer_Hour { get; set; } = string.Empty;
        public static string[] Printer_PrinterName { get; set; } = [];
        public static string[] Printer_PrinterStatus { get; set; } = [];
        public static string Printer_ElapsedTime { get; set; } = string.Empty;


        // Miscellaneous variables for program functions
        public static string Version = "0.9.1"; // Version of the program
        public static bool IsAutoRun { get; set; } = false;
        public static bool WordIsInstalled { get; set; } = false;
        public static int Timeout { get; set; } = 1000; // Timeout in milliseconds
        public static string PrinterList { get; set; } = string.Empty; // Printer list for the printer test, string separated by semicolons
        public static string DrivesList { get; set; } = string.Empty; // Drive letter list for the network storage rights test, string separated by semicolons
        public static string Target { get; set; } = "http://www.google.com"; // Target for the internet test.
        public static List<bool> TestToRun { get; set; } = new List<bool> { true, true, true, true }; // Array of booleans to determine which tests to run, 0: Internet, 1: Network Storage Rights, 2: Office Version/Rights, 3: Printer

        public static void ResetVariables()
        {
            General_TotalTests = 0;
            General_TotalSuccess = 0;
            General_Resume = string.Empty;
            InternetConnexion_Hour = string.Empty;
            InternetConnexion_HTMLStatut = string.Empty;
            InternetConnexion_ElapsedTime = string.Empty;
            NetworkStorageRights_Hour = string.Empty;
            NetworkStorageRights_DiskLetter = [];
            NetworkStorageRights_CheminUNC = [];
            NetworkStorageRights_Serveur = [];
            NetworkStorageRights_ShareName = [];
            NetworkStorageRights_ElapsedTime = string.Empty;
            OfficeVersion_Hour = string.Empty;
            OfficeVersion_OfficeVersion = string.Empty;
            OfficeVersion_OfficePath = string.Empty;
            OfficeVersion_OfficeCulture = string.Empty;
            OfficeVersion_OfficeExcludedApps = string.Empty;
            OfficeVersion_OfficeLastUpdateStatus = string.Empty;
            OfficeVersion_ElapsedTime = string.Empty;
            OfficeRights_Hour = string.Empty;
            OfficeRights_Write = string.Empty;
            OfficeRights_Read = string.Empty;
            OfficeRights_Delete = string.Empty;
            OfficeRights_Save = string.Empty;
            OfficeRights_Create = string.Empty;
            OfficeRights_FolderTested = Path.GetTempPath();
            OfficeRights_ElapsedTime = string.Empty;
            PrinterSystemList = [];
            Printer_Hour = string.Empty;
            Printer_PrinterName = [];
            Printer_PrinterStatus = [];
            Printer_ElapsedTime = string.Empty;
        }
    }
}
