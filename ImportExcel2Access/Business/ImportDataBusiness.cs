using ImportExcel2Access.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace ImportExcel2Access.Business
{
    public class ImportDataBusiness
    {
        private OleDbTransaction accessTransac = null;
        private OleDbConnection accessConnection = null;


        public ImportDataBusiness()
        {
            accessConnection = new OleDbConnection(GetConnection.GetAccessConnectionString);
            SqlHelps.CommandTimeOut = 120;
        }

        /// <summary>
        /// Execute import datatable to access
        /// </summary>
        /// <param name="importData"></param>
        public void ImportExecute(DataTable importData)
        {
            string excelUpdateQuery = string.Empty;
            string accessQuery = string.Empty;
            DataTable tbBaseGroup = new DataTable();
            try
            {
                accessConnection = new OleDbConnection(GetConnection.GetAccessConnectionString);

                if (accessConnection.State == ConnectionState.Closed)
                    accessConnection.Open();

                accessTransac = accessConnection.BeginTransaction();

                accessQuery = string.Format("select top 1 {0}.No from {0}  order by {0}.No desc", Constant.ACCESS_TABLE_NAME);
                var maxNoTable = SqlHelps.ExecuteDataset(GetConnection.GetAccessConnectionString, CommandType.Text, accessQuery).Tables[0];

                int maxNo = 0;

                if (maxNoTable.Rows.Count == 0)
                    maxNo = 1;
                else
                    maxNo = int.Parse(maxNoTable.Rows[0][0].ToString()) + 1;

                string acceesInsertQuery = string.Empty;
                string columns = string.Empty;
                string values = string.Empty;
                string columName = string.Empty;
                int rowCount = 1;
                object value = null;
                int colCount = 0;

                //Get base group
                tbBaseGroup = AccessGetBaseGroup();

                foreach (DataRow row in importData.Rows)
                {
                    // insert to access                   
                    if (row["No"] == null)
                        continue;

                    colCount++;

                    Parameters parameters = new Parameters();
                    parameters.Add(new Parameter()
                    {
                        Name = "@No",
                        Values = maxNo
                    });

                    columns = string.Format("[No],");

                    //Mapping column excel  with access
                    foreach (DataColumn col in importData.Columns)
                    {
                        columName = col.ColumnName;

                        //Ignor column No, 取込
                        if (columName.Equals(Constant.ROW_INDEX_HEADER_TEXT) || columName.Equals("No") || columName.Equals("取込"))
                            continue;

                        //Reset value
                        value = null;

                        //Change column name in excel to name in access
                        switch (columName)
                        {
                            case "記入グループ":
                                if (tbBaseGroup != null && tbBaseGroup.Rows.Count > 0)
                                {
                                    value = tbBaseGroup.Rows.Cast<DataRow>().Where(cl => cl.Field<string>("フィールド1").Equals(row[col].ToString())).Select(r => r.Field<int>("ID")).FirstOrDefault();
                                }
                                break;
                            case "発注時音源":
                                //Change to column name access
                                columName = "発注時音源(正規盤)";
                                if (row[col] != null && row[col].ToString().Equals("正規盤"))
                                    value = true;
                                else
                                    value = false;
                                break;
                            case "発注時歌詞":
                                columName = "発注時歌詞(正規盤)";
                                if (row[col] != null && row[col].ToString().Equals("正規盤"))
                                    value = true;
                                else
                                    value = false;
                                break;
                            case "修正済":
                                if (row[col] != null && row[col].ToString().Equals("済"))
                                    value = true;
                                else
                                    value = false;

                                break;
                            default:
                                value = row[col];
                                break;
                        }

                        //Set value for parameters
                        parameters.Add(new Parameter()
                        {
                            Name = string.Format("@{0}", col.ColumnName),
                            Values = value == null ? DBNull.Value : value
                        });

                        columns += string.Format("[{0}],", columName);
                    }

                    //Removed last comma
                    columns = columns.Remove(columns.Length - 1);
                    //Get parameters name
                    values = string.Join(",", parameters.GetParameters.Select(r => r.Name));

                    //Generate insert query
                    acceesInsertQuery = string.Format("INSERT INTO {0}({1}) VALUES({2})", Constant.ACCESS_TABLE_NAME, columns, values);

                    //Execute insert to transaction
                    SqlHelps.ExecuteNonQuery(accessTransac, CommandType.Text, acceesInsertQuery, parameters);

                    rowCount++;

                    maxNo++;
                }

                // Commit to access
                accessTransac.Commit();
                accessConnection.Close();

                //Excel update column status
                ExcelHelps.UpdateExcelColumn(GetConnection.GetExcelPath, Constant.SHEET_NAME, importData,  Constant.COLUMN_IMPORT_STATUS_HEADER_TEXT, 1);
            }
            catch (Exception ex)
            {
                //Excel reset column status if commit has exception
                accessTransac.Rollback();
                accessConnection.Close();
                throw ex;
            }
        }


        /// <summary>
        /// Get Base group
        /// </summary>
        /// <returns></returns>
        private DataTable AccessGetBaseGroup()
        {
            try
            {
                string query = "select * from  記入グループ";
                return SqlHelps.ExecuteDataset(GetConnection.GetAccessConnectionString, CommandType.Text, query).Tables[0];
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Get data to import from excel
        /// </summary>
        /// <param name="colums">Columns valid</param>
        /// <returns></returns>
        public DataTable GetDataFilter(DataTable dataTable, List<string> colums)
        {
            try
            {
                DataTable dt = new DataTable();

                if (dataTable == null)
                    return null;

                //Removed row = 1
                var tmpTb = dataTable.Rows.Cast<DataRow>().Where(r => r.Field<object>("No") != null && r.Field<object>("修正済") != null && r.Field<object>("修正済").ToString().Equals("済") && (r.Field<object>("取込") == null || r.Field<object>("取込") != null && !r.Field<object>("取込").ToString().Equals("1")));

                if (tmpTb.Count() > 0)
                    dt = tmpTb.CopyToDataTable();

                //Get current column in table
                List<string> currentColumns = Utils.GetColumns(dt);

                // Get columns not valid 
                var excepColumns = currentColumns.Except(colums).ToList();

                //Removed column not valid 
                foreach (var col in excepColumns)
                {
                    if (col.Equals(Constant.ROW_INDEX_HEADER_TEXT)) 
                        continue;
                    dt.Columns.Remove(col);
                }
                dt.AcceptChanges();
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
