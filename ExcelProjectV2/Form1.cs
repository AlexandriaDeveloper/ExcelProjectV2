using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProjectV2
{
    using ExcelProjectV2.NPOI;
    using System.IO;

 
    public partial class Form1 : Form
    {   string tempPath = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void قتاToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void تحديثقواعدالبياناتToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private NPOIFactory npoi = null;
        private void btn_importExcel_Click(object sender, EventArgs e)
        {
            OFD_Excel.RestoreDirectory = true;
            OFD_Excel.Filter = "Excel  files (*.xls;*.xlsx)|*.xls;*.xlsx";
            this.lb_sheets.Items.Clear();

            if (OFD_Excel.ShowDialog() == DialogResult.OK)
            {
                this.lbl_Path.Text = OFD_Excel.FileName;
                npoi = new NPOIFactory(this.lbl_Path.Text);             
            }
            foreach (var sheetName in this.npoi.GetSheetsName())
            {
                this.lb_sheets.Items.Add(sheetName);
            }
        }
        private DataTable dt;
        private void button1_Click(object sender, EventArgs e)
        {

            if (this.npoi != null)
            {
                List<string> sheetsname = new List<string>();
                List<int> sheetsindex = new List<int>();
                this.dg_data.DataSource = null;
                
                foreach (var selectedItem in this.lb_sheets.SelectedItems)
                {
                    sheetsname.Add(selectedItem.ToString());
                    sheetsindex.Add(lb_sheets.Items.IndexOf(selectedItem));
                     
                }
                string paymentType = string.Empty;
                if (this.rb_Atm.Checked)
                {
                    paymentType = "2-اخرى بطاقات حكومية";
                }
                else
                {
                    paymentType = "3-مرتب تحويلات بنكية";
                }



                npoi = new NPOIFactory(this.lbl_Path.Text);
                //TODO Create TempFile
                IntropExcel temp = new IntropExcel();
             tempPath=temp.CreateTempFile(lbl_Path.Text,sheetsindex);



                if (npoi != null)
                {
                    npoi = new NPOIFactory(tempPath);

                }

                this.dt = npoi.ReadExcel(sheetsname, int.Parse(this.txt_RowNum.Text) - 1, paymentType);
                this.dg_data.DataSource = this.dt;
            }
            else
            {
                MessageBox.Show("تأكد من أختيار ملف");
            }

        }

        private void btn_Save_Click(object sender, EventArgs e)
        {
            if (this.dt != null)
            {
                this.npoi.GeneratFiles();
                if (File.Exists(tempPath))
                {
                    File.Delete(tempPath);
                }
                MessageBox.Show("تم بنجاح");
            }
        }

        private void txt_RowNum_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }
}
