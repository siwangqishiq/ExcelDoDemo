using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ReadXls
{
    class Program
    {
        static void Main(string[] args)
        {
            DataSet tableSet = ToDataTable("D:\\ll.xlsx");
            if (tableSet == null)
            {
                Console.WriteLine("读取Excel文件错误!");
            }
            else
            {
                Console.WriteLine("读取成功!");
            }

            //Console.WriteLine(tableSet.GetXml());


            Console.WriteLine(tableSet.Tables[0].Columns[1]);

            for (int i = 0; i < tableSet.Tables.Count; i++)
            {
                Console.WriteLine("分表" + tableSet.Tables[i].TableName + "...");
                wirteToExcel(tableSet.Tables[i], "E:\\" + tableSet.Tables[i].TableName + ".xlsx");
            }//end for i
            Console.WriteLine("OK!");
            Console.ReadKey();
        }


        public static System.Data.DataTable GetExcelTableName(string fileName)
        {
            try
            {
                if (System.IO.File.Exists(fileName))//判断文件是否存在
                {
                    string strConn = "Provider=Microsoft.Ace.OleDb.12.0;" + "Data Source=" +
                            fileName + ";" + "Extended Properties='Excel 12.0; HDR=Yes; IMEX=1'";
                    OleDbConnection conn = new OleDbConnection(strConn);
                    conn.Open();
                    System.Data.DataTable table = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    conn.Close();
                    return table;
                }//end if
                return null;
            }
            catch
            {
                return null;
            }
        }

        /**
         * 
         * */
        public static System.Data.DataTable ExcelToDataTable(string excelName)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" +
                excelName + ";" + "Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";
            string sql = "select * from [sheet1$]";
            DataSet dataSet = new DataSet();
            OleDbConnection conn = new OleDbConnection(sql);
            conn.Open();

            OleDbDataAdapter adapter = new OleDbDataAdapter(sql, conn);
            adapter.Fill(dataSet, "");
            conn.Close();


            return dataSet.Tables[""];
        }

        /// <summary>
        /// 读取excel文件
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static DataSet ToDataTable(string filePath)
        {
            string connStr = "";
            string fileType = System.IO.Path.GetExtension(filePath);
            if (string.IsNullOrEmpty(fileType))
            {
                return null;
            }

            if (fileType == ".xls")
            {
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX\"";
            }
            else
            {
                connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            }

            string sql_F = "Select * FROM [{0}]";

            OleDbConnection conn = null;
            OleDbDataAdapter da = null;
            System.Data.DataTable dtSheetName = null;

            DataSet ds = new DataSet();

            try
            {
                conn = new OleDbConnection(connStr);
                conn.Open();
                dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                HashSet<string> hashSet = new HashSet<string>();
                for (int i = 0; i < dtSheetName.Rows.Count; i++)
                {
                    string sheetName = (string)dtSheetName.Rows[i]["TABLE_NAME"];
                    Console.WriteLine("--->" + sheetName);
                    //if (sheetName.Contains("$") && !sheetName.Replace("'", "").EndsWith("$")) 
                    //{
                        //Console.WriteLine("不要--->"+sheetName);
                        //continue;
                    //}
                    if (sheetName.Contains("FilterDatabase")) 
                    {
                        continue;
                    }
                    

                    if (!hashSet.Contains(sheetName))
                    {
                        hashSet.Add(sheetName);
                    }
                }//end for i


                da = new OleDbDataAdapter();
                foreach(string sheetName in hashSet)
                {
                    Console.WriteLine(String.Format(sql_F, sheetName));
                    da.SelectCommand = new OleDbCommand(String.Format(sql_F, sheetName), conn);
                    DataSet dsItem = new DataSet();
                    
                    da.Fill(dsItem, sheetName);

                    ds.Tables.Add(dsItem.Tables[0].Copy());
                }//end for each
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    da.Dispose();
                    conn.Dispose();
                }
            }

            return ds;
        }

        /*
         * public static bool SaveDataTableToExcel(System.Data.DataTable excelTable, string filePath)
        {
            Microsoft.Office.Interop.Excel.Application app =
                new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                app.Visible = false;
                Workbook wBook = app.Workbooks.Add(true);
                Worksheet wSheet = wBook.Worksheets[1] as Worksheet;
                //Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
                if (excelTable.Rows.Count > 0)
                {
                    int row = 0;
                    row = excelTable.Rows.Count;
                    int col = excelTable.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        for (int j = 0; j < col; j++)
                        {
                            string str = excelTable.Rows[i][j].ToString();
                            wSheet.Cells[i + 2, j + 1] = str;
                        }
                    }
                }
                int size = excelTable.Columns.Count; //写入列名
                for (int i = 0; i < size; i++)
                {
                    wSheet.Cells[1, 1 + i] = excelTable.Columns[i].ColumnName;
                }
                //设置禁止弹出保存和覆盖的询问提示框
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                //保存工作簿
                wBook.Save();
                //保存excel文件
                app.Save(filePath);
                app.SaveWorkspace(filePath);
                app.Quit();
                app = null;
                return true;
            }
            catch (Exception err)
            {
                MessageBox.Show("导出Excel出错！错误原因：" + err.Message, "提示信息",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            finally
            {
            }
        }
         * **/
        public static void wirteToExcel(System.Data.DataTable excelTable,string filePath)
        {
            object objOpt = Missing.Value;
            Application excel = new Application();
            excel.Visible = false;//excel程序不可见
            _Workbook wBook = excel.Workbooks.Add(objOpt);
            _Worksheet wSheet = (_Worksheet)wBook.ActiveSheet;
            wSheet.Visible = XlSheetVisibility.xlSheetVisible;

            int rowIndex = 1;
            int colIndex = 0;

            foreach(DataColumn col in excelTable.Columns)
            {
                colIndex++;
                excel.Cells[1, colIndex] = col.ColumnName;
            }//end for
            foreach (DataRow row in excelTable.Rows)
            {
                rowIndex++;
                colIndex = 0;
                foreach (DataColumn col in excelTable.Columns)
                {
                    colIndex++;
                    excel.Cells[rowIndex, colIndex] = row[col.ColumnName].ToString();
                }
            }

            wBook.SaveAs(filePath, objOpt, null, null, false, false, XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
            wBook.Close(false, objOpt, objOpt);
            excel.Quit();
        }

    }//end class

}//end namespace
