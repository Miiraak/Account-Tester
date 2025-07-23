using BlobPE;

namespace AccountTester
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Dictionary<string, int> defaultData = new Dictionary<string, int>
            {
                { "Langage", 8 },
                { "BaseExtension", 8 },
                { "Timeout", 8 },
                { "AutoExport", 8 },
                { "Autorun", 8 },
                { "PrinterList", 128},
                { "DrivesList", 128 },
                { "Target", 128 }
            };
                Blob.CheckForUpdates(args, defaultData);

                if (args.Length > 0 && args[0] == "--autorun")
                    Variables.IsAutoRun = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }
}