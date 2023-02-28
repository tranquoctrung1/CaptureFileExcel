using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;

namespace CaptureExcelFile.Actions
{
    public class HideRowAction
    {
        public void HideRow(string path, List<int> indexs, string productid)
        {
            try
            {
                Workbook workbook = new Workbook(path);

                // Accessing the first worksheet in the Excel file
                Worksheet worksheet = workbook.Worksheets[2];

                foreach (int i in indexs)
                {
                    worksheet.Cells.HideRow(i + 1);
                }

                // Saving the modified Excel file
                workbook.Save(System.AppDomain.CurrentDomain.BaseDirectory + @"./" + productid + "_output.xlsx");

                workbook.CloseAccessCache(AccessCacheOptions.All);
                workbook.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            // Instantiate a workbook
           
            //foreach (Process process in Process.GetProcessesByName("Excel"))
            //{
            //    process.Kill();
            //}
        }


        public void HideRowImportGoods(string path, List<int> indexs, string productid)
        {
            try
            {
                // Instantiate a workbook
                Workbook workbook = new Workbook(path);

                // Accessing the first worksheet in the Excel file
                Worksheet worksheet = workbook.Worksheets[3];

                foreach (int i in indexs)
                {
                    worksheet.Cells.HideRow(i + 1);
                }

                // Saving the modified Excel file
                workbook.Save(System.AppDomain.CurrentDomain.BaseDirectory + @"./" + productid + "_output2.xlsx");

                workbook.CloseAccessCache(AccessCacheOptions.All);
                workbook.Dispose();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           

            //foreach (Process process in Process.GetProcessesByName("Excel"))
            //{
            //    process.Kill();
            //}
        }

        public void HideRowExportGoods(string path, List<int> indexs, string productid)
        {
            try
            {
                // Instantiate a workbook
                Workbook workbook = new Workbook(path);

                // Accessing the first worksheet in the Excel file
                Worksheet worksheet = workbook.Worksheets[4];

                foreach (int i in indexs)
                {
                    worksheet.Cells.HideRow(i + 1);
                }

                // Saving the modified Excel file
                workbook.Save(System.AppDomain.CurrentDomain.BaseDirectory + @"./" + productid + "_output3.xlsx");

                workbook.CloseAccessCache(AccessCacheOptions.All);
                workbook.Dispose();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            

            //foreach (Process process in Process.GetProcessesByName("Excel"))
            //{
            //    process.Kill();
            //}
        }
    }
}
