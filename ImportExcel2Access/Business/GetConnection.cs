using ImportExcel2Access.Util;
using System;
using System.Data.OleDb;
using static ImportExcel2Access.Util.LogUtil;

namespace ImportExcel2Access.Business
{
    public class GetConnection
    {
        private static string database_connection_string_default = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source={0}; Persist Security Info=False;";
        private static string excel_connection_string_default = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";//Update no need IMEX
        private static string excel_connection_string_reading_default = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=NO;TypeGuessRows=0;IMEX=1;ImportMixedTypes=Text\"";//HDR=YES; with sheet table has header; IMEX=1 for column mix data TypeGuessRows=100;
        private static FileConfig fileConfig;
        private static string databasePath = string.Empty;
        private static string excel_path = string.Empty;

        private static string excel_connection_string = string.Empty;
        private static string excel_connection_string_reading = string.Empty;
        private static string database_connection_string = string.Empty;


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

                database_connection_string = string.Format(database_connection_string_default, databasePath);

                using (OleDbConnection conn = new OleDbConnection(database_connection_string))
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
        public static bool CheckExcelConnection(string excelPath, string tableName)
        {
            try
            {
                excel_connection_string_reading = string.Format(excel_connection_string_reading_default, excelPath);
                excel_connection_string = string.Format(excel_connection_string_default, excelPath);
                excel_path = excelPath;
                using (OleDbConnection conn = new OleDbConnection(excel_connection_string))
                {
                    conn.Open();
                    OleDbCommand command = new OleDbCommand();

                    command.Connection = conn;
                    command.CommandType = System.Data.CommandType.Text;
                    command.CommandText = string.Format("select * from {0}", tableName);
                    command.ExecuteNonQuery();
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

        public static string GetExcelConnectionStringReading
        {
            get
            {
                return excel_connection_string_reading;
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
