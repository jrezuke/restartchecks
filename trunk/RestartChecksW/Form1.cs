using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

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
                var orderedFis = fis.OrderBy(f => f.CreationTime).Reverse();

                foreach (var fi in orderedFis)
                {
                    if (fi.Extension != ".xlsm")
                        continue;
                    
                    listBox2.Items.Add(fi);
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

            string fileName = ((FileInfo)listBox2.SelectedItem).FullName;
            Process.Start("Excel.exe", fileName);
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }

    //public class ExcelProcess
    //{
    //    public string Name { get; set; }
    //}
}
