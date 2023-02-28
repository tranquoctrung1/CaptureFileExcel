using CaptureExcelFile.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CaptureExcelFile.Actions
{
    public class CreateDescriptionAction
    {
        public ContentModel CreateDescriptionFromTotalSheet(System.Data.DataTable table, string listProductId)
        {
            ContentModel el = new ContentModel();

            List<string> splitProductId = listProductId.Split(new char[] { ',' }, StringSplitOptions.None).ToList();

            if(table.Rows.Count > 0)
            {
                foreach(DataRow row in table.Rows)
                {
                    if (row[0] != null)
                    {
                        if (row[0].ToString() != "")
                        {
                            string find = splitProductId.Find(x => x.ToLower() == row[0].ToString().ToLower());

                            //if (listProductId.Contains(row[0].ToString()) == true)
                            if(find != null && find != "")
                            {

                                if(el.ProductName == null || el.ProductName == "")
                                {
                                    el.ProductName = row[1].ToString(); 
                                }
                                if(el.OldDescription == null || el.OldDescription == "")
                                {
                                    el.OldDescription = row[17].ToString();
                                }

                                if (row[3] != null)
                                {
                                    if (row[3].ToString() != "" && row[3].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.StockStartMonth += int.Parse(row[3].ToString());
                                        }
                                        catch(Exception ex)
                                        {
                                            MessageBox.Show($"Tồn đầu kỳ {ex.Message}");
                                        }
                                    }
                                }
                                if (row[4] != null)
                                {
                                    if (row[4].ToString() != "" && row[4].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.ImportVK += int.Parse(row[4].ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Nhập vk {ex.Message}");
                                        }
                                    }
                                }
                                if (row[5] != null)
                                {
                                    if (row[5].ToString() != "" && row[5].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.ImportNCQ += int.Parse(row[5].ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Nhập NCQ {ex.Message}");
                                        }
                                    }
                                }
                                if (row[6] != null)
                                {
                                    if (row[6].ToString() != "" && row[6].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.NhatNam += int.Parse(row[6].ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Nhất Nam {ex.Message}");
                                        }
                                    }
                                }
                                if (row[7] != null)
                                {
                                    if (row[7].ToString() != "" && row[7].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.ImportSW += int.Parse(row[7].ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Nhập SW {ex.Message}");
                                        }
                                    }
                                }
                                if (row[8] != null)
                                {
                                    if (row[8].ToString() != "" && row[8].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.ImportCLK += int.Parse(row[8].ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Nhập CLK {ex.Message}");
                                        }
                                    }
                                }

                                if (row[9] != null)
                                {
                                    if (row[9].ToString() != "" && row[9].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.ImportTL += int.Parse(row[9].ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Nhập TL {ex.Message}");
                                        }
                                    }
                                }
                                if (row[10] != null)
                                {
                                    if (row[10].ToString() != "" && row[10].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.ChangeShield += int.Parse(row[10].ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Đổi vỏ {ex.Message}");
                                        }
                                    }
                                }

                                if (row[11] != null)
                                {
                                    if (row[11].ToString() != "" && row[11].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.ExportSold += int.Parse(row[11].ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Xuat ban {ex.Message}");
                                        }
                                    }
                                }
                                if (row[12] != null)
                                {
                                    if (row[12].ToString() != "" && row[12].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.ExportTransport += int.Parse(row[12].ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Xuat dieu chuyen {ex.Message}");
                                        }
                                    }
                                }
                                if (row[13] != null)
                                {
                                    if (row[13].ToString() != "" && row[13].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.StockEndMonth += int.Parse(row[13].ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Ton cuoi ky {ex.Message}");
                                        }
                                    }
                                }
                                if (row[14] != null)
                                {
                                    if (row[14].ToString() != "" && row[14].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.MiniStock += int.Parse(row[14].ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Kho Mini {ex.Message}");
                                        }
                                    }
                                }
                                if (row[16] != null)
                                {
                                    if (row[16].ToString() != "" && row[16].ToString() != "-")
                                    {
                                        try
                                        {
                                            el.Different += int.Parse(row[16].ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"Kho Mini {ex.Message}");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return el;
        }
    }
}
