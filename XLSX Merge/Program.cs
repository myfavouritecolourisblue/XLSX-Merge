using System.Diagnostics;

namespace XLSX_Merge
{
    internal static class Program
    {
        private static bool NOUI = true;

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            // TODO: If --noui command line argument is set, do not start the UI
            foreach(string arg in args) {
                if (arg.Equals("--noui"))
                    NOUI = true;

            }
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }
    }
}