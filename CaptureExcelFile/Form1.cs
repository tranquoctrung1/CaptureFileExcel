﻿using CaptureExcelFile.Actions;
using CaptureExcelFile.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.DesignerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace CaptureExcelFile
{
    public partial class Form1 : Form
    {
        string pathFile;
        Microsoft.Office.Interop.Excel.Application xl;
        ReadFileExcelAction readFileExcelAction;
        HideRowAction hideRowAction;
        CaptureExcelAction captureExcelAction;
        ValidationExcelFileAction validationExcelFileAction;
        CreateDescriptionAction createDescriptionAction;
        CreateFolderImageAction createFolderImageAction;
        CreateTotalRowAction createTotalRowAction;
        int lengthToCaptureTotalSheet = 2;
        int lengthToCaptureImportGoodsSheet = 2;
        int lengthToCaptureExportGoodsSheet = 2;
        string pathToSaveFileImage;
        bool checkSplit;
        public Form1()
        {
            InitializeComponent();
            readFileExcelAction = new ReadFileExcelAction();
            hideRowAction = new HideRowAction();
            captureExcelAction = new CaptureExcelAction();
            validationExcelFileAction = new ValidationExcelFileAction();
            createDescriptionAction = new CreateDescriptionAction();
            checkSplit = false;
            createFolderImageAction = new CreateFolderImageAction();
            createTotalRowAction = new CreateTotalRowAction();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            if (xl == null)
            {
                MessageBox.Show("Chưa chọn file excel!!!");
                return;
            }
            else
            {
                if(pathToSaveFileImage != "")
                {
                    if(txtDate.Value.ToString() != "")
                    {
                        string productids = txtProductId.Text;

                        if (productids != "")
                        {
                            Cursor.Current = Cursors.WaitCursor;
                            // merge productid
                            if(checkSplit == false)
                            {
                                string folder = $"{DateTime.Now.Day}_{DateTime.Now.Month}_{DateTime.Now.Year}";
                                string lastfixFile = $"{DateTime.Now.Hour}_{DateTime.Now.Minute}_{DateTime.Now.Second}";
                                string prefixFile = productids.Replace(',', '_');

                                // filter file excel
                                System.Data.DataTable table = readFileExcelAction.ReadTotalSheet(pathFile);
                                List<List<int>> listIndex = validationExcelFileAction.GetListIndexToHideTotalSheet(productids, table);


                                if (listIndex.Count > 0)
                                {
                                    hideRowAction.HideRow(pathFile, listIndex[0],prefixFile);
                                }

                                if (listIndex[1].Count > 0)
                                {
                                    lengthToCaptureTotalSheet = listIndex[1][listIndex[1].Count - 1];
                                }

                                System.Data.DataTable tableImportGoods = readFileExcelAction.ReadImportGoodsSheet(pathFile);
                                List<List<int>> listIndexImportGoods = validationExcelFileAction.GetListIndexToHideImportGoodsSheet(productids, tableImportGoods);


                                if (listIndexImportGoods.Count > 0)
                                {
                                    hideRowAction.HideRowImportGoods(pathFile, listIndexImportGoods[0], prefixFile);
                                }

                                if (listIndexImportGoods[1].Count > 0)
                                {
                                    lengthToCaptureImportGoodsSheet = listIndexImportGoods[1][listIndexImportGoods[1].Count - 1];
                                }

                                int rowTotalImportGoods =  createTotalRowAction.CreateTotalRowImportGoodsSheet(AppDomain.CurrentDomain.BaseDirectory +prefixFile +"_output2.xlsx",tableImportGoods, productids);
                                lengthToCaptureImportGoodsSheet = rowTotalImportGoods;

                                

                                createFolderImageAction.CreateFolderImage(pathToSaveFileImage, folder);

                                // capture excel file
                                captureExcelAction.CaptureExcelWithTotalSheet(lengthToCaptureTotalSheet, pathToSaveFileImage, folder, prefixFile, lastfixFile);
                                captureExcelAction.CaptureExcelWithImportGoodsSheet(lengthToCaptureImportGoodsSheet, pathToSaveFileImage, folder, prefixFile, lastfixFile);

                                // create file and capture with each productid
                                string[] productIdSplit = productids.Split(new char[] { ',' }, StringSplitOptions.None);
                                System.Data.DataTable tableExportGoods = readFileExcelAction.ReadExportGoodsSheet(pathFile);

                                foreach(string productid in productIdSplit)
                                {
                                    List<List<int>> listIndexExportGoods = validationExcelFileAction.GetListIndexToHideExportGoodsSheet(productid, tableExportGoods);

                                    if (listIndexExportGoods.Count > 0)
                                    {
                                        hideRowAction.HideRowExportGoods(pathFile, listIndexExportGoods[0], productid);
                                    }

                                    if (listIndexExportGoods[1].Count > 0)
                                    {
                                        lengthToCaptureExportGoodsSheet = listIndexExportGoods[1][listIndexExportGoods[1].Count - 1];
                                    }
                                    prefixFile = productid;

                                    int rowTotalExportGoods =  createTotalRowAction.CreateTotalRowExportGoodsSheet(AppDomain.CurrentDomain.BaseDirectory + prefixFile +"_output3.xlsx", tableExportGoods, productid);
                                    lengthToCaptureExportGoodsSheet = rowTotalExportGoods;

                                    captureExcelAction.CaptureExcelWithExportGoodsSheet(lengthToCaptureExportGoodsSheet, pathToSaveFileImage, folder, prefixFile, lastfixFile);
                                }

                                // create description for total sheet
                                ContentModel content = createDescriptionAction.CreateDescriptionFromTotalSheet(table, productids);

                                string importvk = "";
                                string importncq = "";
                                string nhatnam = "";
                                string importsw = "";
                                string importclk = "";
                                string importtl = "";
                                string changeshield = "";
                                string exportsold = "";
                                string exporttransport = "";
                                string diff = "";
                                if (content.ImportVK != 0)
                                {
                                    importvk = $", NHẬP VK {content.ImportVK} ";
                                }

                                if (content.ImportNCQ != 0)
                                {
                                    importncq = $",NHẬP NCQ {content.ImportNCQ} ";
                                }

                                if (content.NhatNam != 0)
                                {
                                    nhatnam = $",NHẤT NAM {content.NhatNam} ";
                                }
                                if (content.ImportCLK != 0)
                                {
                                    importclk = $", NHẬP CLK {content.ImportCLK} ";
                                }
                                if (content.ImportSW != 0)
                                {
                                    importsw = $", NHẬP SW {content.ImportSW} ";
                                }
                                if (content.ImportTL != 0)
                                {
                                    importtl = $",NHẬP TL {content.ImportTL} ";
                                }
                                if (content.ChangeShield != 0)
                                {
                                    changeshield = $", ĐỔI VỎ {content.ChangeShield} ";
                                }
                                if (content.ExportSold != 0)
                                {
                                    exportsold = $", XUẤT BÁN {content.ExportSold} ";
                                }
                                if (content.ExportTransport != 0)
                                {
                                    exporttransport = $", XUẤT ĐIỀU CHUYỂN {content.ExportTransport} ";
                                }

                                if (content.Different >= 0)
                                {
                                    diff = "THỪA";
                                }
                                else
                                {
                                    diff = "THIẾU";
                                }

                                string description = $"NGÀY {txtDate.Value.Day}/{txtDate.Value.Month} {content.ProductName} TỒN ĐẦU {content.StockStartMonth} {importvk} {importncq} {nhatnam} {importsw} {importclk} {importtl} {changeshield} {exportsold} {exporttransport} = {content.StockEndMonth} KHO TỒN {content.MiniStock} {diff} {content.Different} ( {content.OldDescription} )";

                                txtDescription.Text = description;
                            }
                            else
                            {
                                string[] productidSplit = productids.Split(new char[] { ',' }, StringSplitOptions.None);

                                System.Data.DataTable table = readFileExcelAction.ReadTotalSheet(pathFile);
                                System.Data.DataTable tableImportGoods = readFileExcelAction.ReadImportGoodsSheet(pathFile);
                                System.Data.DataTable tableExportGoods = readFileExcelAction.ReadExportGoodsSheet(pathFile);

                                string totalDescription = "";

                                string folder = $"{DateTime.Now.Day}_{DateTime.Now.Month}_{DateTime.Now.Year}";
                                string lastfixFile = $"{DateTime.Now.Hour}_{DateTime.Now.Minute}_{DateTime.Now.Second}";

                                createFolderImageAction.CreateFolderImage(pathToSaveFileImage, folder);

                                foreach (string productid in productidSplit)
                                {

                                    string prefixFile = productid;

                                    // filter file excel
                                    
                                    List<List<int>> listIndex = validationExcelFileAction.GetListIndexToHideTotalSheet(productid, table);

                                    if (listIndex.Count > 0)
                                    {
                                        hideRowAction.HideRow(pathFile, listIndex[0], prefixFile);
                                    }

                                    if (listIndex[1].Count > 0)
                                    {
                                        lengthToCaptureTotalSheet = listIndex[1][listIndex[1].Count - 1];
                                    }
                                    
                                    List<List<int>> listIndexImportGoods = validationExcelFileAction.GetListIndexToHideImportGoodsSheet(productid, tableImportGoods);

                                    if (listIndexImportGoods.Count > 0)
                                    {
                                        hideRowAction.HideRowImportGoods(pathFile, listIndexImportGoods[0], prefixFile);
                                    }

                                    if (listIndexImportGoods[1].Count > 0)
                                    {
                                        lengthToCaptureImportGoodsSheet = listIndexImportGoods[1][listIndexImportGoods[1].Count - 1];
                                    }

                                    int rowTotalImportGoods = createTotalRowAction.CreateTotalRowImportGoodsSheet(AppDomain.CurrentDomain.BaseDirectory + prefixFile + "_output2.xlsx", tableImportGoods, productids);
                                    lengthToCaptureImportGoodsSheet = rowTotalImportGoods;

                                    List<List<int>> listIndexExportGoods = validationExcelFileAction.GetListIndexToHideExportGoodsSheet(productid, tableExportGoods);

                                    if (listIndexExportGoods.Count > 0)
                                    {
                                        hideRowAction.HideRowExportGoods(pathFile, listIndexExportGoods[0], productid);
                                    }

                                    if (listIndexExportGoods[1].Count > 0)
                                    {
                                        lengthToCaptureExportGoodsSheet = listIndexExportGoods[1][listIndexExportGoods[1].Count - 1];
                                    }
                                    prefixFile = productid;

                                    int rowTotalExportGoods = createTotalRowAction.CreateTotalRowExportGoodsSheet(AppDomain.CurrentDomain.BaseDirectory + prefixFile + "_output3.xlsx", tableExportGoods, productid);
                                    lengthToCaptureExportGoodsSheet = rowTotalExportGoods;

                                    // capture excel file
                                    captureExcelAction.CaptureExcelWithTotalSheet(lengthToCaptureTotalSheet, pathToSaveFileImage, folder, prefixFile, lastfixFile);
                                    captureExcelAction.CaptureExcelWithImportGoodsSheet(lengthToCaptureImportGoodsSheet, pathToSaveFileImage, folder, prefixFile, lastfixFile);
                                    captureExcelAction.CaptureExcelWithExportGoodsSheet(lengthToCaptureExportGoodsSheet, pathToSaveFileImage, folder, prefixFile, lastfixFile);

                                    // create description for total sheet
                                    ContentModel content = createDescriptionAction.CreateDescriptionFromTotalSheet(table, productid);

                                    string importvk = "";
                                    string importncq = "";
                                    string nhatnam = "";
                                    string importsw = "";
                                    string importclk = "";
                                    string importtl = "";
                                    string changeshield = "";
                                    string exportsold = "";
                                    string exporttransport = "";
                                    string diff = "";
                                    if (content.ImportVK != 0)
                                    {
                                        importvk = $", NHẬP VK {content.ImportVK} ";
                                    }

                                    if (content.ImportNCQ != 0)
                                    {
                                        importncq = $",NHẬP NCQ {content.ImportNCQ} ";
                                    }

                                    if (content.NhatNam != 0)
                                    {
                                        nhatnam = $",NHẤT NAM {content.NhatNam} ";
                                    }
                                    if (content.ImportCLK != 0)
                                    {
                                        importclk = $", NHẬP CLK {content.ImportCLK} ";
                                    }
                                    if (content.ImportSW != 0)
                                    {
                                        importsw = $", NHẬP SW {content.ImportSW} ";
                                    }
                                    if (content.ImportTL != 0)
                                    {
                                        importtl = $",NHẬP TL {content.ImportTL} ";
                                    }
                                    if (content.ChangeShield != 0)
                                    {
                                        changeshield = $", ĐỔI VỎ {content.ChangeShield} ";
                                    }
                                    if (content.ExportSold != 0)
                                    {
                                        exportsold = $", XUẤT BÁN {content.ExportSold} ";
                                    }
                                    if (content.ExportTransport != 0)
                                    {
                                        exporttransport = $", XUẤT ĐIỀU CHUYỂN {content.ExportTransport} ";
                                    }

                                    if (content.Different >= 0)
                                    {
                                        diff = "THỪA";
                                    }
                                    else
                                    {
                                        diff = "THIẾU";
                                    }

                                    totalDescription += $"- {productid}: NGÀY {txtDate.Value.Day}/{txtDate.Value.Month} {content.ProductName} TỒN ĐẦU {content.StockStartMonth} {importvk} {importncq} {nhatnam} {importsw} {importclk} {importtl} {changeshield} {exportsold} {exporttransport} = {content.StockEndMonth} KHO TỒN {content.MiniStock} {diff} {content.Different} ( {content.OldDescription} ) {Environment.NewLine}";
                                }

                                txtDescription.Text = totalDescription;
                            }

                            Cursor.Current = Cursors.Default;
                        }
                        else
                        {
                            MessageBox.Show("Chưa có mã hàng hóa!!!");
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Chưa chọn ngày!!");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Chưa chọn nơi lưu ảnh!!!");
                    return;
                }

            }
        }


        private void btn_browserFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "Excel Files|*.xls;*.xlsx";

            of.Title = "Chọn file excel";

            if (of.ShowDialog() == DialogResult.OK)
            {
                pathFile = of.FileName;

                xl = new Microsoft.Office.Interop.Excel.Application();

            }
        }

        private void btnChoosePathToSaveImage_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog of = new FolderBrowserDialog();

            if (of.ShowDialog() == DialogResult.OK)
            {
                pathToSaveFileImage = of.SelectedPath;

            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;

            if(now.Year >= 2023 && now.Month >= 5 && now.Day >= 1)
            {
                lbTitle.Text = "PHẦN MỀM ĐÃ HẾT HẠN. CẦN PHẢI KÍCH HOẠT ĐỂ DÙNG LẠI PHẦN MỀM!!!";
                lbTitle.Location = new System.Drawing.Point(10, 29);
                lbTitle.ForeColor = Color.Red;
                btnChoosePathToSaveImage.Enabled = false;
                btn_browserFile.Enabled = false;
                button2.Enabled = false;
                txtDate.Enabled = false;
                txtProductId.Enabled = false;
                txtDescription.Enabled = false;
            }
        }

        private void ckSplit_CheckedChanged(object sender, EventArgs e)
        {
            checkSplit = ckSplit.Checked;
        }

    }
}