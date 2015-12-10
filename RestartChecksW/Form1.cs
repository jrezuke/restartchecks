using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace RestartChecksW
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            //listBox1.DisplayMember = "MainWindowTitle";
            InitializeComponent();
            this.Icon = RestartChecksW.Properties.Resources.restart;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitializeForm();
        }

        private void InitializeForm()
        {
            listBox2.DisplayMember = "Name";
            listBox2.ValueMember = "FullName";
            KillAllExcelProcesses();
            GetAllChecksFiles();
        }

        private void GetAllChecksFiles()
        {
            const string path = @"C:\Halfpint";

            if (Directory.Exists(path))
            {
                var di = new DirectoryInfo(path);

                FileInfo[] fis = di.GetFiles();
                var orderedFis = fis.OrderBy(f => f.LastWriteTime).Reverse();

                foreach (var fi in orderedFis)
                {
                    if (fi.Extension != ".xlsm")
                        continue;
                    if (fi.Name.StartsWith("~"))
                        continue;
                    listBox2.Items.Add(fi.Name + " (" + fi.LastWriteTime +")");
                }
            }
        }

        private void KillAllExcelProcesses()
        {
            var excelProcesses = Process.GetProcessesByName(processName: "Excel");
            foreach (var excelProcess in excelProcesses)
            {
                excelProcess.Kill();
            }
        }
       
        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedItem == null)
            {
                MessageBox.Show("Select a subject to restart!");
                return;
            }

            var selected = listBox2.SelectedItem.ToString();
            //var splits = selected.Split(' ');
            var pos = selected.IndexOf(' ');
            var filename = selected.Substring(0, pos);
            var datePart = selected.Substring(pos + 2).Replace(")","");
            
            var date = DateTime.Parse(datePart);
            string fullName = GetFilenameFromSelected(filename, date);

            Process.Start("Excel.exe", fullName);
            Application.Exit();
        }

        private string GetFilenameFromSelected(string fileName, DateTime date)
        {
            const string path = @"C:\Halfpint";

            if (Directory.Exists(path))
            {
                var di = new DirectoryInfo(path);

                FileInfo[] fis = di.GetFiles();
                var orderedFis = fis.OrderBy(f => f.LastWriteTime).Reverse();

                foreach (var fi in orderedFis)
                {
                    if (fi.Extension != ".xlsm")
                        continue;
                    if (fi.Name.StartsWith("~"))
                        continue;
                    if (fi.Name == fileName )
                    {
                        var diff = fi.LastWriteTime - date;
                        if(diff.Days == 0 && diff.Hours ==0 && diff.Minutes==0 && diff.Seconds == 0) 
                            return fi.FullName;
                    }
                }
            }
            return "";
        }

    }

    
}
