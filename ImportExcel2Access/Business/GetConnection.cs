using ImportExcel2Access.Util;
using System;
using System.Data.OleDb;
using System.Text;
using static ImportExcel2Access.Util.LogUtil;

namespace ImportExcel2Access.Business
{
    public class GetConnection
    {
        private static string database_connection_string = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Persist Security Info=False;";
        private static string excel_connection_string = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"";//HDR=YES; with sheet table has header; IMEX=1 for column mix data
        private static FileConfig fileConfig;
        private static string databasePath = string.Empty;
        private static string excel_path = string.Empty;

        /// <summary>
        /// Check connection string
        /// </summary>
        /// <returns>TRUE: Connected - FLASE: Fails</returns>
        public static bool CheckAccessConnection()
        {
            try
            {
                fileConfig = new FileConfig(Constant.FILE_CONFIG);
                databasePath = fileConfig.Read(Constant.AccessKey, Constant.ACCESS_SECSION);

                database_connection_string = string.Format(database_connection_string, databasePath);

                using (OleDbConnection conn = new OleDbConnection(database_connection_string))
                {
                    conn.Open();
                    conn.Close();
                }

                return true;
            }
            catch(Exception ex)
            {
                ErrorEntity error = new ErrorEntity()
                {
                    ErrorMessage = ex.Message,
                    FunctionName = "CheckAccessConnection",
                    FilePath = Constant.LOG_FILE_PATH
                };
                LogUtil.Write(error);
                return false;
            }
        }

        /// <summary>
        /// Get connection string
        /// </summary>
        public static string GetAccessConnectionString
        {
            get
            {
                return database_connection_string;
            }
        }

        public static string GetAccessPath
        {
            get
            {
                return databasePath;
            }
        }

        /// <summary>
        /// Check excel connection
        /// </summary>
        /// <param name="excelPath"></param>
        /// <returns></returns>
        public static bool CheckExcelConnection(string excelPath)
        {
            try
            {
                excel_connection_string = string.Format(excel_connection_string, excelPath);               
                excel_path = excelPath;
                using (OleDbConnection conn = new OleDbConnection(excel_connection_string))
                {
                    conn.Open();
                    conn.Close();
                }
                return true;
            }
            catch (Exception ex)
            {
                ErrorEntity error = new ErrorEntity()
                {
                    ErrorMessage = ex.Message,
                    FunctionName = "CheckExcelConnection",
                    FilePath = Constant.LOG_FILE_PATH
                };
                LogUtil.Write(error);
                return false;
            }
        }

        /// <summary>
        /// Get excel connection string
        /// </summary>
        public static string GetExcelConnectionString
        {
            get
            {
                return excel_connection_string;
            }
        }

        public static string GetExcelPath
        {
            get
            {
                return excel_path;
            }
        }
    }
}
