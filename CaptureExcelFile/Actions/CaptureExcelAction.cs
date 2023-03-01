using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaptureExcelFile.Actions
{
    public class CaptureExcelAction
    {
        public void CaptureExcelWithTotalSheet(int length, string pathToSaveFileImage,string folder, string prefixFile, string lastfixFile)
        {
            Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
            if (xl == null)
            {
                MessageBox.Show("Chưa chọn được file excel!!");
                return;
            }
            else
            {
                try
                {
                   Excel.Workbook wb = xl.Workbooks.Open(System.AppDomain.CurrentDomain.BaseDirectory + prefixFile + "_output.xlsx");
                   Excel.Range r = wb.Sheets[3].Range[$"A1:R{length + 2}"];
                   r.CopyPicture(Excel.XlPictureAppearance.xlScreen,
                                   Excel.XlCopyPictureFormat.xlBitmap);

                    if (Clipboard.GetDataObject() != null)
                    {
                        IDataObject data = Clipboard.GetDataObject();

                        if (data.GetDataPresent(DataFormats.Bitmap))
                        {
                            Image image = (Image)data.GetData(DataFormats.Bitmap, true);
                            image.Save(pathToSaveFileImage + "\\" + folder+ "\\" +prefixFile +"_tonghop_"+lastfixFile+".jpg",
                                System.Drawing.Imaging.ImageFormat.Jpeg);
                        }
                        else
                        {
                            MessageBox.Show("No image in Clipboard !!");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Clipboard Empty !!");
                    }

                    wb.Close();
                    xl.Quit();
                    xl = null;
                }
                catch (Exception  ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            //foreach (Process process in Process.GetProcessesByName("Excel"))
            //{
            //    process.Kill();
            //}
        }

        public void CaptureExcelWithImportGoodsSheet(int length, string pathToSaveFileImage,string folder, string prefixFile, string lastfixFile)
        {
            Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
            if (xl == null)
            {
                MessageBox.Show("Chưa chọn được file excel!!");
                return;
            }
            else
            {
                try
                {
                    Excel.Workbook wb = xl.Workbooks.Open(System.AppDomain.CurrentDomain.BaseDirectory +prefixFile + "_output2.xlsx");
                    Excel.Range r = wb.Sheets[4].Range[$"A1:R{length}"];
                    r.CopyPicture(Excel.XlPictureAppearance.xlScreen,
                                   Excel.XlCopyPictureFormat.xlBitmap);

                    if (Clipboard.GetDataObject() != null)
                    {
                        IDataObject data = Clipboard.GetDataObject();

                        if (data.GetDataPresent(DataFormats.Bitmap))
                        {
                            Image image = (Image)data.GetData(DataFormats.Bitmap, true);
                            image.Save(pathToSaveFileImage + "\\" + folder + "\\" + prefixFile + "_nhaphang_" + lastfixFile + ".jpg",
                               System.Drawing.Imaging.ImageFormat.Jpeg);
                        }
                        else
                        {
                            MessageBox.Show("No image in Clipboard !!");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Clipboard Empty !!");
                    }

                    wb.Close();
                    xl.Quit();
                    xl = null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            //foreach (Process process in Process.GetProcessesByName("Excel"))
            //{
            //    process.Kill();
            //}
        }

        public void CaptureExcelWithExportGoodsSheet(int length, string pathToSaveFileImage, string folder, string prefixFile, string lastfixFile)
        {
            Microsoft.Office.Interop.Excel.Application xl = new Microsoft.Office.Interop.Excel.Application();
            if (xl == null)
            {
                MessageBox.Show("Chưa chọn được file excel!!");
                return;
            }
            else
            {
                try
                {
                    Excel.Workbook wb = xl.Workbooks.Open(System.AppDomain.CurrentDomain.BaseDirectory + prefixFile +"_output3.xlsx");
                    Excel.Range r = wb.Sheets[5].Range[$"A1:L{length}"];
                    r.CopyPicture(Excel.XlPictureAppearance.xlScreen,
                                   Excel.XlCopyPictureFormat.xlBitmap);

                    if (Clipboard.GetDataObject() != null)
                    {
                        IDataObject data = Clipboard.GetDataObject();

                        if (data.GetDataPresent(DataFormats.Bitmap))
                        {
                            Image image = (Image)data.GetData(DataFormats.Bitmap, true);
                            image.Save(pathToSaveFileImage + "\\" + folder + "\\" + prefixFile + "_xuathang_" + lastfixFile + ".jpg",
                               System.Drawing.Imaging.ImageFormat.Jpeg);
                        }
                        else
                        {
                            MessageBox.Show("No image in Clipboard !!");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Clipboard Empty !!");
                    }

                    wb.Close();
                    xl.Quit();
                    xl = null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            //foreach (Process process in Process.GetProcessesByName("Excel"))
            //{
            //    process.Kill();
            //}
        }
    }
}
