using System.Diagnostics;


namespace RestartChecks
{
    class Program
    {
        static void Main(string[] args)
        {
            var localByName = Process.GetProcessesByName(processName: "Excel");
            
            var title = localByName[0].MainWindowTitle;
            
            var pos = title.IndexOf("- ", System.StringComparison.Ordinal);
            
            var fileName = title.Substring(pos + 2);
            
            var fileInfo = "c:/halfpint/" + fileName;
            
            foreach (var process in localByName)
                process.Kill();


            Process.Start("Excel.exe", fileInfo);
        }
    }
}
