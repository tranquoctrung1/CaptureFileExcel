using CaptureExcelFile.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaptureExcelFile.Actions
{
    public class CalculateTotalAction
    {
        public TotalImportGoodsModel CalculateTotalImportGoods(System.Data.DataTable table, string productid)
        {
            TotalImportGoodsModel el = new TotalImportGoodsModel();

            if(table.Rows.Count > 0)
            {
                foreach(DataRow row in table.Rows)
                {
                    if (row[3] != null)
                    {
                        if (row[3].ToString() != "" && row[3].ToString() != "-")
                        {
                            if (productid.ToLower().Contains(row[3].ToString().ToLower()))
                            {
                                if (row[7] != null)
                                {
                                    if (row[7].ToString() != "" && row[7].ToString() != "-")
                                    {
                                        el.ImportVK += int.Parse(row[7].ToString());
                                    }
                                }
                                if (row[8] != null)
                                {
                                    if (row[8].ToString() != "" && row[8].ToString() != "-")
                                    {
                                        el.ImportNCQ += int.Parse(row[8].ToString());
                                    }
                                }
                                if (row[9] != null)
                                {
                                    if (row[9].ToString() != "" && row[9].ToString() != "-")
                                    {
                                        el.NhatNam += int.Parse(row[9].ToString());
                                    }
                                }

                                if (row[10] != null)
                                {
                                    if (row[10].ToString() != "" && row[10].ToString() != "-")
                                    {
                                        el.ImportSW += int.Parse(row[10].ToString());
                                    }
                                }

                                if (row[11] != null)
                                {
                                    if (row[11].ToString() != "" && row[11].ToString() != "-")
                                    {
                                        el.ChangeShield += int.Parse(row[11].ToString());
                                    }
                                }

                                if (row[12] != null)
                                {
                                    if (row[12].ToString() != "" && row[12].ToString() != "-")
                                    {
                                        el.ImportCLK += int.Parse(row[12].ToString());
                                    }
                                }

                                if (row[13] != null)
                                {
                                    if (row[13].ToString() != "" && row[13].ToString() != "-")
                                    {
                                        el.ImportTL += int.Parse(row[13].ToString());
                                    }
                                }
                            }
                        }
                    }
                    
                }
            }


            return el;
        }

        public TotalExportGoodsModel CalculateTotalExportGoods(System.Data.DataTable table, string productid)
        {
            TotalExportGoodsModel el = new TotalExportGoodsModel();

            if(table.Rows.Count > 0) { 
                foreach(DataRow row in table.Rows) {

                    if (row[3] != null)
                    {
                        if (row[3].ToString() != "" && row[3].ToString() != "-")
                        {
                            if (productid.ToLower().Contains(row[3].ToString().ToLower()))
                            {
                                if (row[6] != null)
                                {
                                    if (row[6].ToString() != "" && row[6].ToString() != "-")
                                    {
                                        el.Amout += int.Parse(row[6].ToString());
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
