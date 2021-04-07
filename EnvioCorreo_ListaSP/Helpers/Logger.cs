using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using static System.Configuration.ConfigurationManager;

namespace EnvioCorreo_ListaSP.Helpers
{
    public class Logger
    {
        public static void WriteLine(string message)
        {
            File.AppendAllText($"{AppSettings["RutaLog"]}\\LogER.txt", $@"{DateTime.Now:dd/MM/yyyy HH:mm:ss} => {message}{Environment.NewLine}");
            Console.WriteLine(message);
        }
    }
}
