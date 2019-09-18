using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Net;
using System.Windows.Forms;

namespace BoschTool
{
    public partial class Ribbon1
    {
        public bool LoginState { get; set; } = false;
        public Worksheet sht;
        public string WorkPath;
        public string FileUploadPath;
        public string FileSavePath;
        public WebPage page;
        public Form1 fm;
        public CookieContainer cookieContainer;

        public void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            sht = Globals.ThisWorkbook.Sheets[1];
            WorkPath = AppDomain.CurrentDomain.BaseDirectory;
            cookieContainer = new CookieContainer();
            page = new WebPage(cookieContainer, "https://sgpftsn2.bosch.com.sg");
        }

        public void LoadLoginForm()
        {
            fm = new Form1(page);
            Dictionary<string, string> arrInputs = new Dictionary<string, string>();
            fm.pictureBox1.Image = page.Captcha_IMG(ref arrInputs);
            fm.arrInputs = arrInputs;
            fm.ShowDialog();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            DialogResult result = fd.ShowDialog();
            if (result == DialogResult.OK)
            {
                int lastRow = sht.Cells[sht.Rows.Count, 1].End[XlDirection.xlUp].Row;
                if (lastRow > 1)
                    sht.Range["a2:e" + lastRow].Value = "";
                FileUploadPath = fd.SelectedPath;
                string[] files = Directory.GetFiles(FileUploadPath, "*.xlsx", System.IO.SearchOption.TopDirectoryOnly);
                for (int i = 0; i < files.Length; i++)
                {
                    sht.Cells[i + 2, 1] = Path.GetFileName(files[i]);
                }
                sht.Columns[1].AutoFit();
                MessageBox.Show("Task finished.", "BOSCH", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }            
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            FolderBrowserDialog fd = new FolderBrowserDialog();
            DialogResult fdResult = fd.ShowDialog();
            if (fdResult == DialogResult.OK)
            {
                FileSavePath = fd.SelectedPath;
                if (page.GetLoginState())
                {
                    int lastRow = sht.Cells[sht.Rows.Count, 3].End[XlDirection.xlUp].Row;
                    if (lastRow <= 1) return;
                    object[,] names = sht.Range["a2:e" + lastRow].Value;
                    
                    for (int i = 0; i < names.GetLength(0); i++)
                    {
                        string shipmentName = names[i + 1, 3].ToString().Trim();
                        if (!String.IsNullOrEmpty(shipmentName))
                        {
                            string result = page.DownloadCCS(shipmentName, FileSavePath);
                            sht.Cells[i + 2, 5] = result;
                        }
                        else
                        {
                            sht.Cells[i + 2, 5] = "Shipment Name is mandatory field.";
                        }
                    }
                    MessageBox.Show("Task finished.", "BOSCH", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    LoginState = false;
                    MessageBox.Show("Please login website.", "BOSCH", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void Button3_Click(object sender, RibbonControlEventArgs e)
        {
            LoadLoginForm();
        }

        private void Button6_Click(object sender, RibbonControlEventArgs e)
        {
            if (page.GetLoginState())
            {
                int lastRow = sht.Cells[sht.Rows.Count, 1].End[XlDirection.xlUp].Row;
                if (lastRow <= 1) return;
                object[,] names = sht.Range["a2:e" + lastRow].Value;
                
                for (int i = 0; i < names.GetLength(0); i++)
                {
                    string fullName = Path.Combine(FileUploadPath, names[i + 1, 1].ToString().Trim());
                    if (File.Exists(fullName))
                    {
                        string result = page.UploadFile(fullName);
                        sht.Cells[i + 2, 5] = result;
                    }
                    else
                    {
                        sht.Cells[i + 2, 5] = "File not found.";
                    }
                }
                MessageBox.Show("Task finished.", "BOSCH", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                LoginState = false;
                MessageBox.Show("Please login website.", "BOSCH", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Button4_Click(object sender, RibbonControlEventArgs e)
        {
            if (page.GetLoginState())
            {
                int lastRow = sht.Cells[sht.Rows.Count, 2].End[XlDirection.xlUp].Row;
                if (lastRow <= 1) return;
                object[,] names = sht.Range["a2:e" + lastRow].Value;

                for (int i = 0; i < names.GetLength(0); i++)
                {
                    string invoiceNumber = names[i + 1, 2].ToString().Trim();
                    string shipmentName = names[i + 1, 3].ToString().Trim();
                    if (!String.IsNullOrEmpty(shipmentName) && !String.IsNullOrEmpty(invoiceNumber))
                    {
                        string result = page.AssignCode(invoiceNumber, shipmentName);
                        sht.Cells[i + 2, 5] = result;
                    }
                    else
                    {
                        sht.Cells[i + 2, 5] = "Shipment Name and Invoice Number are mandatory fields.";
                    }
                }
                MessageBox.Show("Task finished.", "BOSCH", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                LoginState = false;
                MessageBox.Show("Please login website.", "BOSCH", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Button5_Click(object sender, RibbonControlEventArgs e)
        {
            if (page.GetLoginState())
            {
                int lastRow = sht.Cells[sht.Rows.Count, 3].End[XlDirection.xlUp].Row;
                if (lastRow <= 1) return;
                object[,] names = sht.Range["a2:e" + lastRow].Value;

                for (int i = 0; i < names.GetLength(0); i++)
                {
                    string tradeType = names[i + 1, 4].ToString().Trim();
                    string shipmentName = names[i + 1, 3].ToString().Trim();
                    if (!String.IsNullOrEmpty(shipmentName) && !String.IsNullOrEmpty(tradeType))
                    {
                        string result = page.GenerateCCS(shipmentName, tradeType);
                        sht.Cells[i + 2, 5] = result;
                    }
                    else
                    {
                        sht.Cells[i + 2, 5] = "Shipment Name and Trade Type are mandatory fields.";
                    }
                }
                MessageBox.Show("Task finished.", "BOSCH", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                LoginState = false;
                MessageBox.Show("Please login website.", "BOSCH", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
