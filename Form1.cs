using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Virus_Checker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ObservableCollection<File> filelist = new ObservableCollection<File>();
            filelist.Add(File.addpath("filepath"));
            filelist[0].MD5_hash = "MD5_hash";
            filelist[0].SHA1_hash = "SHA1_hash";
            filelist[0].virustotal_url = "virustotal_url";

            this.fileBindingSource.DataSource = filelist;
        }

        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] dlagFilePathArray = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            for (int i = 0; i < dlagFilePathArray.Length; i++)
            {
                fileBindingSource.Add(File.addpath(dlagFilePathArray[i]));
            }
        }

        private void button1_MouseClick(object sender, MouseEventArgs e)
        {
            ObservableCollection<File> filelist = this.fileBindingSource.DataSource as ObservableCollection<File>;
            for (int i = 1; i < filelist.Count; i++)
            {
                filelist[i].Calc_MD5();
                filelist[i].Calc_SHA1();
                filelist[i].Calc_SHA256();
            }
            OutputExcel.ExcelWriter(filelist);
            this.Close();
        }
    }
}
