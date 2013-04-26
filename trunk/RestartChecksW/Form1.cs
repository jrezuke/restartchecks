using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RestartChecksW
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            //listBox1.DisplayMember = "MainWindowTitle";
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitializeForm();
        }

        private void InitializeForm()
        {
            listBox2.DisplayMember = "Name";
            listBox2.ValueMember = "FullName";
            GetAllExcelProcesses();
            GetAllChecksFiles();
        }

        private void GetAllChecksFiles()
        {
            var path = ConfigurationManager.AppSettings["ChecksPath"].ToString();
            
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

        private void GetAllExcelProcesses()
        {
            listBox1.Items.Clear();
            var excelProcesses = Process.GetProcessesByName(processName: "Excel");
            foreach (var excelProcess in excelProcesses)
            {
                listBox1.Items.Add(excelProcess.MainWindowTitle);
            }
        }

        private void btnEndProcess_Click(object sender, EventArgs e)
        {
            var excelProcesses = Process.GetProcessesByName(processName: "Excel");
            foreach (var excelProcess in excelProcesses)
            {
                excelProcess.Kill();
                listBox1.Items.Remove(excelProcess.MainWindowTitle);
            }
            GetAllExcelProcesses();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedValue == null)
            {
                MessageBox.Show("Select a CHECKS file to restart!");
                return;
            }

            //string fileName = listBox2.SelectedValue.ToString();
            //Process.Start("Excel.exe", fileName);
        }
    }

    public class ExcelProcess
    {
        public string Name { get; set; }
    }
}
