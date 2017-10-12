using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Virus_Checker
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
            System.Windows.Forms.Application.Run(new Form1());
        }
    }

    public class File
    {
        private string filepathValue = String.Empty;

        public File()
        { }

        public static File addpath(string v)
        {
            File tmp = new File();
            tmp.filepathValue = v;
            return tmp;
        }

        public string filepath
        {
            get
            {
                return this.filepathValue;
            }

            set
            {
                this.filepathValue = value;
            }
        }

        public string MD5_hash;
        public string SHA1_hash;
        public string SHA256_hash;
        public string virustotal_url;

        public void Calc_MD5()
        {
            System.IO.FileStream fs = new System.IO.FileStream(
                filepath,
                System.IO.FileMode.Open,
                System.IO.FileAccess.Read,
                System.IO.FileShare.Read);

            System.Security.Cryptography.MD5CryptoServiceProvider md5
                = new System.Security.Cryptography.MD5CryptoServiceProvider();

            byte[] bs = md5.ComputeHash(fs);

            md5.Clear();
            fs.Close();

            MD5_hash = BitConverter.ToString(bs).ToLower().Replace("-", "");
        }

        public void Calc_SHA1()
        {
            System.IO.FileStream fs = new System.IO.FileStream(
                filepath,
                System.IO.FileMode.Open,
                System.IO.FileAccess.Read,
                System.IO.FileShare.Read);

            System.Security.Cryptography.SHA1CryptoServiceProvider sha1
                = new System.Security.Cryptography.SHA1CryptoServiceProvider();

            byte[] bs = sha1.ComputeHash(fs);

            sha1.Clear();
            fs.Close();

            SHA1_hash = BitConverter.ToString(bs).ToLower().Replace("-", "");
        }

        public void Calc_SHA256()
        {
            System.IO.FileStream fs = new System.IO.FileStream(
                filepath,
                System.IO.FileMode.Open,
                System.IO.FileAccess.Read,
                System.IO.FileShare.Read);

            System.Security.Cryptography.SHA256CryptoServiceProvider sha256
                = new System.Security.Cryptography.SHA256CryptoServiceProvider();

            byte[] bs = sha256.ComputeHash(fs);

            sha256.Clear();
            fs.Close();

            SHA256_hash = BitConverter.ToString(bs).ToLower().Replace("-", "");
            virustotal_url = "https://www.virustotal.com/ja/file/" + SHA256_hash + "/analysis/";
        }
    }

    public static class OutputExcel
    {
        public static void ExcelWriter(ObservableCollection<File> list)
        {
            Excel.Application ExcelApp = null;
            Excel.Workbooks wbs = null;
            Excel.Workbook wb = null;
            Excel.Sheets shs = null;
            Excel.Worksheet ws = null;

            try
            {
                ExcelApp = new Excel.Application();
                wbs = ExcelApp.Workbooks;
                wb = wbs.Add();

                shs = wb.Sheets;
                ws = shs[1];
                ws.Select(Type.Missing);

                ExcelApp.Visible = false;

                for (int i = 1; i <= list.Count; i++)
                {
                    Excel.Range w_rgn = ws.Cells;
                    Excel.Range rgn = w_rgn[i, 1];

                    try
                    {
                        rgn.Value2 = list[i - 1].filepath;
                        rgn = w_rgn[i, 2];
                        rgn.Value2 = list[i - 1].MD5_hash;
                        rgn = w_rgn[i, 3];
                        rgn.Value2 = list[i - 1].SHA1_hash;
                        rgn = w_rgn[i, 4];
                        rgn.Value2 = list[i - 1].virustotal_url;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(w_rgn);
                        Marshal.ReleaseComObject(rgn);
                        w_rgn = null;
                        rgn = null;
                    }
                }

                string name = System.DateTime.Now.ToString("yyyyMMdd_HH-mm-ss");
                wb.SaveAs(System.IO.Directory.GetCurrentDirectory() + "\\" + name + ".xlsx");
                wb.Close(false);
                MessageBox.Show("File Outputed [" + name + ".xlsx]");
                ExcelApp.Quit();
            }
            finally
            {
                Marshal.ReleaseComObject(ws);
                Marshal.ReleaseComObject(shs);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(wbs);
                Marshal.ReleaseComObject(ExcelApp);
                ws = null;
                shs = null;
                wb = null;
                wbs = null;
                ExcelApp = null;

                GC.Collect();
            }

        }
    }
}