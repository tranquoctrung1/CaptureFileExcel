using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CaptureExcelFile.Actions
{
    public class ValidationExcelFileAction
    {
        public List<List<int>> GetListIndexToHideTotalSheet(string listProductId, System.Data.DataTable table)
        {
            List<List<int>> listIndex = new List<List<int>>();

            List<int> list = new List<int>();
            listIndex.Add(list);
            List<int> list2 = new List<int>();
            listIndex.Add(list2);

            if (table.Rows.Count > 0)
            {
                int count = 0;
                foreach(DataRow row in table.Rows)
                {
                    if (count != 0)
                    {
                        if (row[0] != null)
                        {
                            if (row[0].ToString() != "")
                            {
                                if (listProductId.ToLower().Contains(row[0].ToString().ToLower()) == false)
                                {
                                    listIndex[0].Add(count);
                                }
                                else
                                {
                                    listIndex[1].Add(count);
                                }
                            }
                            else
                            {
                                listIndex[0].Add(count);   
                            }
                        }
                        else
                        {
                            listIndex[0].Add(count);
                        }
                       
                    }
                    count += 1;
                }
            }

            return listIndex;
        }


        public List<List<int>> GetListIndexToHideImportGoodsSheet(string listProductId, System.Data.DataTable table)
        {
            List<List<int>> listIndex = new List<List<int>>();

            List<int> list = new List<int>();
            listIndex.Add(list);
            List<int> list2 = new List<int>();
            listIndex.Add(list2);

            if (table.Rows.Count > 0)
            {
                int count = 0;
                foreach (DataRow row in table.Rows)
                {
                    if (count != 0)
                    {
                        if (row[3] != null)
                        {
                            if (row[3].ToString() != "")
                            {
                                if (listProductId.ToLower().Contains(row[3].ToString().ToLower()) == false)
                                {
                                    listIndex[0].Add(count);
                                }
                                else
                                {

                                    listIndex[1].Add(count);
                                }

                            }
                            else
                            {
                                listIndex[0].Add(count);
                               
                            }
                        }
                        else
                        {
                            listIndex[0].Add(count);
                            
                        }

                    }
                    count += 1;
                }
            }

            return listIndex;
        }

        public List<List<int>> GetListIndexToHideExportGoodsSheet(string listProductId, System.Data.DataTable table)
        {
            List<List<int>> listIndex = new List<List<int>>();

            List<int> list = new List<int>();
            listIndex.Add(list);
            List<int> list2 = new List<int>();
            listIndex.Add(list2);

            if (table.Rows.Count > 0)
            {
                int count = 0;
                foreach (DataRow row in table.Rows)
                {
                    if (count != 0)
                    {
                        if (row[3] != null)
                        {
                            if (row[3].ToString() != "")
                            {
                                if (listProductId.ToLower().Contains(row[3].ToString().ToLower()) == false )
                                {
                                    listIndex[0].Add(count);
                                }
                                else
                                {
                                    listIndex[1].Add(count);
                                }

                            }
                            else
                            {
                                listIndex[0].Add(count);
                            }
                        }
                        else
                        {
                            listIndex[0].Add(count);
                        }

                    }
                    count += 1;
                }
            }

            return listIndex;
        }
    }
}
