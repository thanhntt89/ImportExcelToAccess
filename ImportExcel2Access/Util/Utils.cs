using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ImportExcel2Access.Util
{
    public class Utils
    {
        /// <summary>
        /// Get columns of data table
        /// </summary>
        /// <param name="dataTable"></param>
        /// <returns></returns>
        public static List<string> GetColumns(DataTable dataTable)
        {
            return dataTable.Columns.Cast<DataColumn>().Select(col => col.ColumnName).ToList();
        }

        public static bool IsExistColumn(string columnName, List<string> columns)
        {
            var exist = columns.Where(r => r.Equals(columnName)).FirstOrDefault();
            return exist != null;
        }

        /// <summary>
        /// Check file is used 
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns></returns>
        public static bool FileLocked(string FileName)
        {
            FileStream fs = null;

            try
            {
                // NOTE: This doesn't handle situations where file is opened for writing by another process but put into write shared mode, it will not throw an exception and won't show it as write locked
                fs = File.Open(FileName, FileMode.Open, FileAccess.ReadWrite, FileShare.None); // If we can't open file for reading and writing then it's locked by another process for writing
            }
            catch (UnauthorizedAccessException) // https://msdn.microsoft.com/en-us/library/y973b725(v=vs.110).aspx
            {
                // This is because the file is Read-Only and we tried to open in ReadWrite mode, now try to open in Read only mode
                try
                {
                    fs = File.Open(FileName, FileMode.Open, FileAccess.Read, FileShare.None);
                }
                catch (Exception)
                {
                    return true; // This file has been locked, we can't even open it to read
                }
            }
            catch (Exception)
            {
                return true; // This file has been locked
            }
            finally
            {
                if (fs != null)
                    fs.Close();
            }
            return false;
        }
    }
}
