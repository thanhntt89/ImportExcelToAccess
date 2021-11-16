using System;
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

                //Get base group
                tbBaseGroup = GetBaseGroup();

                foreach (DataRow row in importData.Rows)
                {
                    // insert to access                   

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
                        if (columName.Equals("No") || columName.Equals("取込"))
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
                            Values = value
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
                
                //Excel update column status
                ExcelHelps.UpdateExcelStatusColumn(GetConnection.GetExcelPath, importData, 1);

                // Commit to access
                accessTransac.Commit();
                accessConnection.Close();
            }
            catch (Exception ex)
            {
                //Excel reset column status if commit has exception
              
                ExcelHelps.UpdateExcelStatusColumn(GetConnection.GetExcelPath, importData, string.Empty);

                accessTransac.Rollback();
                accessConnection.Close();
                throw ex;
            }
        }


        /// <summary>
        /// Get base id
        /// </summary>
        /// <returns></returns>
        private DataTable GetBaseId()
        {
            try
            {
                string query = "select * from 拠点ID";
                return SqlHelps.ExecuteDataset(GetConnection.GetAccessConnectionString, CommandType.Text, query).Tables[0];
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Get Base group
        /// </summary>
        /// <returns></returns>
        private DataTable GetBaseGroup()
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
        /// <returns></returns>
        public DataTable GetDataFromExcel()
        {
            try
            {
                string excelSelectQuery = string.Format("select * from {0} where 修正済= '済' and 取込 is null", Constant.EXCEL_TABLE_NAME);
                return SqlHelps.ExecuteDataset(GetConnection.GetExcelConnectionString, CommandType.Text, excelSelectQuery).Tables[0];
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
