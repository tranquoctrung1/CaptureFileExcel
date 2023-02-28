using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CaptureExcelFile.Actions
{
    public class ReadFileExcelAction
    {
        public System.Data.DataTable ReadTotalSheet(string path)
        {
            System.Data.DataTable table = new System.Data.DataTable();
            try
            {
                string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection(connStr);
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [TỔNG HỢP$]", MyConnection);
                MyCommand.TableMappings.Add("Table", "Net-informations.com");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                MyConnection.Close();
                MyConnection.Dispose();
                MyCommand.Dispose();

                if (DtSet.Tables.Count > 0)
                {

                    table = DtSet.Tables[0];
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return table;
        }

        public System.Data.DataTable ReadImportGoodsSheet(string path)
        {

            System.Data.DataTable table = new System.Data.DataTable();
            try
            {
                string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection(connStr);
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [NHẬP HÀNG$]", MyConnection);
                MyCommand.TableMappings.Add("Table", "Net-informations.com");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                MyConnection.Close();
                MyConnection.Dispose();
                MyCommand.Dispose();

                if (DtSet.Tables.Count > 0)
                {

                    table = DtSet.Tables[0];
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return table;
        }

        public System.Data.DataTable ReadExportGoodsSheet(string path)
        {

            System.Data.DataTable table = new System.Data.DataTable();
            try
            {
                string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.DataSet DtSet;
                System.Data.OleDb.OleDbDataAdapter MyCommand;
                MyConnection = new System.Data.OleDb.OleDbConnection(connStr);
                MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [XUẤT HÀNG$]", MyConnection);
                MyCommand.TableMappings.Add("Table", "Net-informations.com");
                DtSet = new System.Data.DataSet();
                MyCommand.Fill(DtSet);
                MyConnection.Close();
                MyConnection.Dispose();
                MyCommand.Dispose();

                if (DtSet.Tables.Count > 0)
                {

                    table = DtSet.Tables[0];
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return table;
        }
    }
}
