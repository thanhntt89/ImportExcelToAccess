using System.Collections.Generic;
using System.Data;
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
    }
}
