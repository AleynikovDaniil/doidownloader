using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Net;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace doidownloader
{
    public delegate void InvokeDelegate();
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            SelfRef = this;
        }

        public static Form1 SelfRef
        {
            get;
            set;
        }

        int x = 2;
        string extention = "";
        bool threading = false;
        string workpath = Environment.CurrentDirectory;

        public void getFileExt(string outFile)
        {
            string firstLine = File.ReadLines(workpath + "\\files\\temp\\" + outFile.Trim() + ".tmp").First();
            if(firstLine.Contains("html"))
            {
                extention = "html";
            }
            else if(firstLine.Contains("pdf") || firstLine.Contains("PDF"))
            {
                extention = "pdf";
            }
        }

        public void download(object doilink)
        {
            int j = x;
            if(threading == true)
            {
                j = x - 1;
            }
            string outfile = dataGridView1.Rows[j - 2].Cells[0].Value.ToString().Replace(".", "_").Replace("/", "_").Replace("\\", "_");
            using (var client = new System.Net.WebClient())
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                try
                {
                    client.DownloadFile(new Uri(doilink.ToString()), workpath + "\\files\\temp\\" + outfile.Trim() + ".tmp");
                    getFileExt(outfile);
                    File.Move(workpath + "\\files\\temp\\" + outfile.Trim() + ".tmp", workpath + "\\files\\" + outfile.Trim() +"."+ extention);
                    dataGridView1.Rows[j - 2].Cells[2].Value = "Загружено";
                    dataGridView1.Rows[j - 2].Cells[3].Value = extention;
                    dataGridView1.Rows[j - 2].Cells[4].Value = outfile + "."+ extention;
                }
                catch(Exception ex)
                {
                    if (ex.Message.Contains("Невозможно создать файл, так как он уже существует"))
                    {
                        dataGridView1.Rows[j - 2].Cells[2].Value = "Уже существует.";
                        dataGridView1.Rows[j - 2].Cells[3].Value = extention;
                        dataGridView1.Rows[j - 2].Cells[4].Value = outfile + "." + extention;
                    }
                    else
                    {
                        dataGridView1.Rows[j - 2].Cells[2].Value = ex.Message;
                        dataGridView1.Rows[j - 2].Cells[3].Value = "Ссылка";
                        dataGridView1.Rows[j - 2].Cells[4].Value = doilink;
                    }
                }
            }            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Thread mainthread = new Thread(new ThreadStart(main));
            mainthread.SetApartmentState(ApartmentState.STA);
            mainthread.Start();
            //main();
        }

        public void main()
        {
            string fldlgres = "";
            Excel.Application excel_app = new Excel.Application();
            OpenFileDialog openfileDialog = new OpenFileDialog();
            openfileDialog.ShowDialog();
            fldlgres = openfileDialog.FileName;
            Thread[] potok = new Thread[200];

            //download();
            Excel.Workbook workbook = excel_app.Workbooks.Open(fldlgres, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[1];
            while (sheet.Cells[x, 1].Text != "")
            {
                potok[x] = new Thread(new ParameterizedThreadStart(download));
                Form1.SelfRef.Invoke(new MethodInvoker(delegate ()
                {
                    Form1.SelfRef.dataGridView1.Rows.Add();
                    Form1.SelfRef.dataGridView1.Rows[x - 2].Cells[0].Value = sheet.Cells[x, 1].Text;
                    Form1.SelfRef.dataGridView1.Rows[x - 2].Cells[1].Value = sheet.Cells[x, 7].Text;
                    Form1.SelfRef.dataGridView1.Rows[x - 2].Cells[2].Value = "В процессе..";
                }));
                if (threading == false)
                {
                    download("https://doi.org/" + dataGridView1.Rows[x - 2].Cells[0].Value.ToString().Trim());
                }
                else
                {
                    potok[x].Start("https://doi.org/" + dataGridView1.Rows[x - 2].Cells[0].Value.ToString().Trim());
                }
                x += 1;
            }
            x = 2;
            excel_app.Quit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string savepath = Environment.CurrentDirectory + "\\" + "output.xls";
            Excel.Application savedexel = new Excel.Application();
            savedexel.Workbooks.Add();
            Excel.Worksheet savedsheet = (Excel.Worksheet)savedexel.ActiveSheet;
            for (int i = 0; i < 5; i++)
            {
                savedsheet.Cells[1, i + 1].Value = dataGridView1.Columns[i].HeaderText;
            }
            for (int i = 0; i < dataGridView1.Rows.Count-1;i++)
            {
                savedsheet.Cells[i + 2, 1] = dataGridView1.Rows[i].Cells[0].Value.ToString();
                savedsheet.Cells[i + 2, 2] = dataGridView1.Rows[i].Cells[1].Value.ToString();
                savedsheet.Cells[i + 2, 3] = dataGridView1.Rows[i].Cells[2].Value.ToString();
                savedsheet.Cells[i + 2, 4] = dataGridView1.Rows[i].Cells[3].Value.ToString();
                savedsheet.Cells[i + 2, 5] = dataGridView1.Rows[i].Cells[4].Value.ToString();
            }
            savedsheet.SaveAs(savepath);
            savedexel.Visible = true;
            //savedexel.Quit();

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            threading = true;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            threading = false;
        }
    }
}
