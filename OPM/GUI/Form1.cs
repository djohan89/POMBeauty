using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OPM.OPMEnginee;
using OPM.ExcelHandler;


namespace OPM
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<Packagelist> oPackagelists = new List<Packagelist>();
        Dictionary<int, string> ListTPO = new Dictionary<int, string>();
        string[] Chosen_Files = null;
        string Chosen_File = string.Empty;

        private void button1_Click(object sender, EventArgs e)
        {
            
            try
            {
                
                //string[] Chosen_File = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                //openFileDialog.InitialDirectory = Chosen_File;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Chosen_Files = openFileDialog1.FileNames;
                }
                if (Chosen_Files.Count() == 0)
                {
                    return;
                }
                textBox1.Text = Chosen_Files[0]+"...";

                //int ret = OpmExcelHandler.fReadPackageListFiles(Chosen_Files, ref oPackagelists);
                /**/
                /**/
            }
            catch (Exception)
            {
                MessageBox.Show("Sorry, Error");
                return;

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                Chosen_File = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog.InitialDirectory = Chosen_File;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Chosen_File = openFileDialog.FileName;
                }
                if (Chosen_File.ToString() == string.Empty)
                {
                    return;
                }

                textBox2.Text = Chosen_File;
                
            }
            catch (Exception)
            {
                MessageBox.Show("Sorry, Error");
                return;

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = this.folderBrowserDialog1.ShowDialog();
                string foldername = "";
                if (result == DialogResult.OK)
                {
                    foldername = this.folderBrowserDialog1.SelectedPath;
                }

                textBox3.Text = foldername;
            }
            catch (Exception)
            {
                MessageBox.Show("Sorry, Error");
                return;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            int ret = OpmExcelHandler.fReadExcelFile(Chosen_File, ref ListTPO);
            if (0 == ret)
            {
                return;
            }
            
            ret = OpmExcelHandler.fReadPackageListFiles(Chosen_Files, ref oPackagelists);
            if (0 == ret)
            {
                return;
            }
            ret = OpmExcelHandler.fWriteExcelfile(ListTPO, oPackagelists, textBox3.Text);
            if (1 == ret)
            {
                return;
            }
            MessageBox.Show("OK Đã Hoàn Thành, Chúc các chị Đẹp vui vẻ");
        }
        private void button6_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
