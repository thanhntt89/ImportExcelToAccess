using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace ImportExcel2Access.Business
{
    public class ExcelHelps
    {
        /// <summary>
        /// Update column status in excel after importing
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="status"></param>
        public static void UpdateExcelStatusColumn(string excelPath, DataTable dataTable, object status)
        {
            try
            {
                string excelUpdateQuery = string.Empty;

                //Update in to excel               
                foreach (DataRow row in dataTable.Rows)
                {
                    excelUpdateQuery = string.Format("update {0} set {0}.取込 = {1} where {0}.No = {2}", Constant.EXCEL_TABLE_NAME, status, row["No"]);

                    SqlHelps.ExecuteNonQuery(GetConnection.GetExcelConnectionString, CommandType.Text, excelUpdateQuery);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Update column in excel files
        /// </summary>
        /// <param name="excelPath">excel path</param>
        /// <param name="sheetName">sheet name</param>
        /// <param name="dataTable">table update data</param>
        /// <param name="columnName">colum update</param>
        /// <param name="value">value update</param>
        public static void UpdateExcelColumn(string excelPath, string sheetName, DataTable dataTable, string columnName, object value)
        {
            try
            {
                //Create COM Objects.
                ExcelApp.Application excelApp = null;
                ExcelApp.Workbook excelBook = null;

                try
                {
                    excelApp = new ExcelApp.Application();
                    excelBook = excelApp.Workbooks.Open(excelPath,
                    Missing.Value, false,
                Missing.Value, Missing.Value,
                Missing.Value, true,
                Missing.Value, Missing.Value,
                Missing.Value, Missing.Value,
                Missing.Value, Missing.Value,
                Missing.Value, Missing.Value);

                    ExcelApp._Worksheet excelSheet = excelBook.Sheets[sheetName];
                    ExcelApp.Range excelRange = excelSheet.UsedRange;
                    
                    object[,] worksheetValuesArray = excelRange.get_Value(Type.Missing);

                    string columHeader = string.Empty;
                    int rows = worksheetValuesArray.GetLength(0);
                    int cols = worksheetValuesArray.GetLength(1);
                    int updateColumnIndex = 0;

                    for (int i = 1; i <= cols; i++)
                    {
                        columHeader = worksheetValuesArray[Constant.START_HEADER_INDEX, i] == null ? string.Format("F{0}", i) : worksheetValuesArray[Constant.START_HEADER_INDEX, i].ToString();
                        if (columHeader.Equals(columnName))
                        {
                            updateColumnIndex = i;
                            break;
                        }
                    }

                    foreach (DataRow row in dataTable.Rows)
                    {
                        worksheetValuesArray[int.Parse(row[Constant.ROW_INDEX_HEADER_TEXT].ToString()), updateColumnIndex] = value;
                    }

                    excelRange.set_Value(Type.Missing, worksheetValuesArray);
                    excelBook.Save();
                }
                catch (Exception ex)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelBook);
                    Marshal.ReleaseComObject(excelApp);
                }
                finally
                {
                    if (excelBook != null)
                    {
                        excelBook.Close(true, Type.Missing, Type.Missing);
                        Marshal.ReleaseComObject(excelBook);
                    }
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }

                    KillSpecificExcelFileProcess(excelPath);
                }
            }
            catch (Exception ex)
            {

            }
        }

        /// <summary>
        /// Reading data from excel
        /// </summary>
        /// <param name="excelPath"></param>
        /// <param name="sheetName"></param>
        /// <param name="rows"></param>
        /// <param name="cols"></param>
        /// <returns></returns>
        public static object[,] GetDataFromExcelToMatrix(string excelPath, string sheetName, ref List<string> Columns)
        {
            //Create COM Objects.
            ExcelApp.Application excelApp = null;
            ExcelApp.Workbook excelBook = null;

            try
            {
                excelApp = new ExcelApp.Application();
                excelBook = excelApp.Workbooks.Open(excelPath,
                Missing.Value, false,
            Missing.Value, Missing.Value,
            Missing.Value, true,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);

                ExcelApp._Worksheet excelSheet = excelBook.Sheets[sheetName];
                ExcelApp.Range excelRange = excelSheet.UsedRange;
                if (excelBook.ReadOnly)
                {
                    return null;
                }

                if (excelApp == null)
                {
                    return null;
                }

                object[,] worksheetValuesArray = excelRange.get_Value(Type.Missing);
                string columName = string.Empty;
                int cols = worksheetValuesArray.GetLength(1);

                for (int i = 1; i <= cols; i++)
                {
                    columName = worksheetValuesArray[Constant.START_HEADER_INDEX, i] == null ? string.Format("F{0}", i) : worksheetValuesArray[Constant.START_HEADER_INDEX, i].ToString();
                    Columns.Add(columName);
                }

                return worksheetValuesArray;
            }
            catch (Exception ex)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelBook);
                Marshal.ReleaseComObject(excelApp);
                return null;
            }
            finally
            {
                if (excelBook != null)
                {
                    excelBook.Close(true, Type.Missing, Type.Missing);
                    Marshal.ReleaseComObject(excelBook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                KillSpecificExcelFileProcess(excelPath);
            }
        }

        public static DataTable ConvertArrayToDataTable(object[,] dataObjectArray)
        {
            DataRow myNewRow = null;
            DataTable myTable = new DataTable("MyDataTable");
            string columName = string.Format(Constant.ROW_INDEX_HEADER_TEXT);
            int columnNoIndex = 0;
            int rows = dataObjectArray.GetLength(0);
            int cols = dataObjectArray.GetLength(1);
            myTable.Columns.Add(columName);
            try
            {
                for (int i = 1; i <= cols; i++)
                {
                    columName = dataObjectArray[Constant.START_HEADER_INDEX, i] == null ? string.Format("F{0}", i) : dataObjectArray[Constant.START_HEADER_INDEX, i].ToString();
                    if (string.IsNullOrWhiteSpace(columName))
                        continue;
                    if (columName.Equals("No")) columnNoIndex = i;

                    myTable.Columns.Add(columName);
                }

                //first row using for heading, start second row for data
                for (int i = Constant.START_DATA_INDEX; i <= rows; i++)
                {
                    myNewRow = myTable.NewRow();

                    if (dataObjectArray[i, columnNoIndex] == null) continue;

                    //RowIndex
                    myNewRow[0] = i;

                    //Get data in columns
                    for (int col = 1; col <= cols; col++)
                    {
                        myNewRow[col] = dataObjectArray[i, col];
                    }

                    myTable.Rows.Add(myNewRow);
                }

            }
            catch (Exception ex)
            {

            }
            return myTable;
        }


        /// <summary>
        /// Reading excel files
        /// </summary>
        /// <param name="excelPath">Excel path</param>
        /// <param name="sheetName">Sheetname</param>
        /// <returns></returns>
        public static DataTable GetDataFromExcel(string excelPath, string sheetName)
        {
            //Create COM Objects.
            ExcelApp.Application excelApp = null;
            ExcelApp.Workbook excelBook = null;

            try
            {
                excelApp = new ExcelApp.Application();
                excelBook = excelApp.Workbooks.Open(excelPath,
                Missing.Value, false,
            Missing.Value, Missing.Value,
            Missing.Value, true,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);

                ExcelApp._Worksheet excelSheet = excelBook.Sheets[sheetName];
                ExcelApp.Range excelRange = excelSheet.UsedRange;
                if (excelBook.ReadOnly)
                {
                    return null;
                }

                DataRow myNewRow;
                DataTable myTable;

                if (excelApp == null)
                {
                    Console.WriteLine("Excel is not installed!!");
                    return null;
                }

                int rows = excelRange.Rows.Count;
                int cols = excelRange.Columns.Count;

                //Set DataTable Name and Columns Name
                myTable = new DataTable("MyDataTable");

                string columName = string.Empty;

                string colsss = string.Empty;
                int columnNoIndex = 0;

                object[,] worksheetValuesArray = excelRange.get_Value(Type.Missing);

                //Stopwatch stopWatch = new Stopwatch();
                //stopWatch.Start();

                //Create columns
                for (int i = 1; i <= cols; i++)
                {
                    columName = worksheetValuesArray[Constant.START_HEADER_INDEX, i] == null ? string.Format("F{0}", i) : worksheetValuesArray[Constant.START_HEADER_INDEX, i].ToString();
                    if (string.IsNullOrWhiteSpace(columName))
                        continue;
                    if (columName.Equals("No")) columnNoIndex = i;

                    myTable.Columns.Add(columName);
                }

                //first row using for heading, start second row for data
                for (int i = Constant.START_DATA_INDEX; i <= rows; i++)
                {
                    myNewRow = myTable.NewRow();

                    if (worksheetValuesArray[i, columnNoIndex] == null) continue;

                    //Get data in columns
                    for (int col = 1; col <= cols; col++)
                    {
                        myNewRow[col - 1] = worksheetValuesArray[i, col];
                    }

                    myTable.Rows.Add(myNewRow);
                }

                //stopWatch.Stop();
                //TimeSpan ts = stopWatch.Elapsed;
                //int s = ts.Seconds;

                //after reading, relaase the excel project              
                return myTable;
            }
            catch (Exception ex)
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelBook);
                Marshal.ReleaseComObject(excelApp);
                return null;
            }
            finally
            {
                if (excelBook != null)
                {
                    excelBook.Close(true, Type.Missing, Type.Missing);
                    Marshal.ReleaseComObject(excelBook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                KillSpecificExcelFileProcess(excelPath);
            }
        }

        private static void KillSpecificExcelFileProcess(string excelFileName)
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                if (process.MainWindowTitle.Equals(string.Empty))
                    process.Kill();
            }
        }
    }
}
