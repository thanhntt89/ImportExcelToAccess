using System;
using System.Data;

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
                    excelUpdateQuery = string.Format("update {0} set 取込 = {1} where {0}.No = {2}", Constant.EXCEL_TABLE_NAME, status, row["No"]);
                    SqlHelps.ExecuteNonQuery(GetConnection.GetExcelConnectionString, CommandType.Text, excelUpdateQuery);
                }
            }
            catch (Exception ex)
            {
                //throw ex;
            }
        }
    }
}
