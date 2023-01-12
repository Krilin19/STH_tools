using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ES_commands
{
    public partial class SyncListUpdater : Form
    {
        public SyncListUpdater()
        {
            InitializeComponent();
        }

        public void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void SyncListUpdater_Load(object sender, EventArgs e)
        {
            
           
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    string Sync_Manager = @"T:\Lopez\Sync_Manager.xlsx";

        //    try
        //    {
        //        using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
        //        {
        //            ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        MessageBox.Show("Excel Sync Manager file not found, try Sync the normal way", "Sync Warning");
                
        //    }
        //    try
        //    {
        //        using (ExcelPackage package = new ExcelPackage(new FileInfo(Sync_Manager)))
        //        {
        //            ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);
        //            var Time_ = DateTime.Now;

                    
        //            if (sheet.Cells[1, 1].Value.ToString() != null)
        //            {
        //                sheet.DeleteRow(1, 1);
        //                sheet.DeleteRow(1, 2);
        //                package.Save();
        //            }
        //        }

        //    }
        //    catch (Exception)
        //    {
                
        //    }
        //}
    }
}
