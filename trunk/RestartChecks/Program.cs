using System;
using System.Diagnostics;


namespace RestartChecks
{
    class Program
    {
        static void Main(string[] args)
        {
            var localByName = Process.GetProcessesByName(processName: "Excel");

            if (localByName.Length > 0)
            {
                var title = localByName[0].MainWindowTitle;

                var pos = title.IndexOf("- ", System.StringComparison.Ordinal);

                var fileName = title.Substring(pos + 2);

                var fileInfo = "c:/halfpint/" + fileName;

                foreach (var process in localByName)
                    process.Kill();


                Process.Start("Excel.exe", fileInfo);
            }
            else
            {
                Console.WriteLine("Excel is not currently running.  Press enter to exit...");
                Console.ReadLine();
            }
        }
    }
}
