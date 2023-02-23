using CaptureExcelFile.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CaptureExcelFile.Actions
{
    public class CreateTotalRowAction
    {
        public int CreateTotalRowImportGoodsSheet(string path, System.Data.DataTable table, string productid)
        {
            CalculateTotalAction calculateTotalAction = new CalculateTotalAction();

            int usedRow = 1;

            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[4];

                TotalImportGoodsModel total = calculateTotalAction.CalculateTotalImportGoods(table, productid);

                usedRow = table.Rows.Count + 2;
                
                xlWorksheet.Cells[usedRow, 1] = "TỔNG";
                xlWorksheet.Range[xlWorksheet.Cells[usedRow, 1], xlWorksheet.Cells[usedRow, 5]].Merge();
                xlWorksheet.Range[xlWorksheet.Cells[usedRow, 1], xlWorksheet.Cells[usedRow, 18]].Interior.Color = ColorTranslator.ToOle(Color.GreenYellow);
                xlWorksheet.Range[xlWorksheet.Cells[usedRow, 1], xlWorksheet.Cells[usedRow, 18]].EntireRow.Font.Bold = true;
                xlWorksheet.Cells[usedRow, 8] = total.ImportVK == 0 ? "-": total.ImportVK.ToString();
                xlWorksheet.Cells[usedRow, 9] = total.ImportNCQ == 0 ? "-" : total.ImportNCQ.ToString();
                xlWorksheet.Cells[usedRow, 10] = total.NhatNam == 0 ? "-" : total.NhatNam.ToString();
                xlWorksheet.Cells[usedRow, 11] = total.ImportSW == 0 ? "-" : total.ImportSW.ToString();
                xlWorksheet.Cells[usedRow, 12] = total.ChangeShield == 0 ? "-" : total.ChangeShield.ToString();
                xlWorksheet.Cells[usedRow, 13] = total.ImportCLK == 0 ? "-": total.ImportCLK.ToString();
                xlWorksheet.Cells[usedRow, 14] = total.ImportTL == 0 ? "-": total.ImportTL.ToString();
                xlWorksheet.Range[xlWorksheet.Cells[usedRow, 8], xlWorksheet.Cells[usedRow, 14]].Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                xlWorksheet.Range[xlWorksheet.Cells[usedRow, 8], xlWorksheet.Cells[usedRow, 14]].Font.Color = ColorTranslator.ToOle(Color.Red);

                xlWorkbook.Close(true);

                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);

                Marshal.ReleaseComObject(xlApp);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                
            }

            return usedRow;
        }

        public int CreateTotalRowExportGoodsSheet(string path, System.Data.DataTable table, string productid)
        {
            CalculateTotalAction calculateTotalAction = new CalculateTotalAction();
            int usedRow = 1;

            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[5];

                TotalExportGoodsModel total = calculateTotalAction.CalculateTotalExportGoods(table, productid);

                usedRow = table.Rows.Count + 2;

                xlWorksheet.Cells[usedRow, 1] = "TỔNG";
                xlWorksheet.Range[xlWorksheet.Cells[usedRow, 1], xlWorksheet.Cells[usedRow, 6]].Merge();
                xlWorksheet.Range[xlWorksheet.Cells[usedRow, 1], xlWorksheet.Cells[usedRow, 6]].Interior.Color = ColorTranslator.ToOle(Color.LightPink);
                xlWorksheet.Range[xlWorksheet.Cells[usedRow, 1], xlWorksheet.Cells[usedRow, 18]].EntireRow.Font.Bold = true;
                xlWorksheet.Cells[usedRow, 7] = total.Amout == 0 ? "-" : total.Amout.ToString();
                xlWorksheet.Range[xlWorksheet.Cells[usedRow, 7], xlWorksheet.Cells[usedRow, 8]].Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                xlWorksheet.Range[xlWorksheet.Cells[usedRow, 7], xlWorksheet.Cells[usedRow, 8]].Font.Color = ColorTranslator.ToOle(Color.Red);

                xlWorkbook.Close(true);

                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);

                Marshal.ReleaseComObject(xlApp);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {

            }
            return usedRow;
        }
    }
}
