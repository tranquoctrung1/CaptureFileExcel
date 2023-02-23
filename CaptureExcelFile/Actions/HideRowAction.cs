using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;

namespace CaptureExcelFile.Actions
{
    public class HideRowAction
    {
        public void HideRow(string path, List<int> indexs, string productid)
        {
            // Instantiate a workbook
            Workbook workbook = new Workbook(path);

            // Accessing the first worksheet in the Excel file
            Worksheet worksheet = workbook.Worksheets[2];


            foreach(int i in indexs)
            {
                worksheet.Cells.HideRow(i+ 1);
            }

            // Saving the modified Excel file
            workbook.Save(System.AppDomain.CurrentDomain.BaseDirectory +@"./" + productid + "_output.xlsx");

            //foreach (Process process in Process.GetProcessesByName("Excel"))
            //{
            //    process.Kill();
            //}
        }


        public void HideRowImportGoods(string path, List<int> indexs, string productid)
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

            //foreach (Process process in Process.GetProcessesByName("Excel"))
            //{
            //    process.Kill();
            //}
        }

        public void HideRowExportGoods(string path, List<int> indexs, string productid)
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
            workbook.Save(System.AppDomain.CurrentDomain.BaseDirectory + @"./" + productid+"_output3.xlsx");

            //foreach (Process process in Process.GetProcessesByName("Excel"))
            //{
            //    process.Kill();
            //}
        }
    }
}
