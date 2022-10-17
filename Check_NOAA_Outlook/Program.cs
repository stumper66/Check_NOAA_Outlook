using System;

namespace Check_NOAA_Outlook
{
    class Program
    {
        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.UnhandledException += new System.UnhandledExceptionEventHandler(HandleTheUnhandled);

            MainProgram MP = new();
            int ExitCode = MP.RunMain();
            Environment.ExitCode = ExitCode;
        }

        private static void HandleTheUnhandled(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = (Exception)e.ExceptionObject;
            string StrCD = Environment.CurrentDirectory;
            if (!StrCD.EndsWith("\\")) StrCD += "\\";

            System.Text.StringBuilder sb = new("Program: Check_NOAA_Outlook\r\n");
            sb.AppendFormat("Version: {0}\r\n", System.Reflection.Assembly.GetExecutingAssembly().GetName().Version);
            Exception ex2 = null;
            if (ex.InnerException != null) ex2 = (Exception)ex.InnerException;

            sb.AppendFormat("exception message: {0}\r\n" +
                "exception object: {1}\r\n" +
                "exception type: {2}\r\n",
                ex.Message, ex, ex.GetType());

            if (ex2 != null)
            {
                sb.AppendFormat("inner exception message: {0}\r\n" +
                "inner exception object: {1}\r\n" +
                "inner exception type: {2}\r\n",
                ex2.Message, ex2, ex2.GetType());
            }

            string file = StrCD + "debug.txt";

            try
            {
                using System.IO.StreamWriter SW = new(file);
                SW.Write(sb.ToString());
                SW.Close();
            }
            catch { }

            Console.WriteLine("Unhandled exception.");
            Console.WriteLine(ex.Message);
            //log.Fatal("Unhandled exception.", ex);
        }
    }
}
