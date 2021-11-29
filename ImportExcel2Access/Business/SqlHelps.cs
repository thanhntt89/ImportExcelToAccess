using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace ImportExcel2Access.Business
{
    public class SqlHelps
    {
        public static int CommandTimeOut { get; set; }
        private static int DefaultTimeOut = 30;

        public static void ExecuteNonQuery(string connectionString, CommandType commandType, string commandText, Parameters parameters = null)
        {
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                OleDbCommand dbCommand = new OleDbCommand();
               
                if (parameters != null)
                {
                    foreach (var par in parameters.GetParameters)
                    {
                        dbCommand.Parameters.Add(new OleDbParameter(par.Name, par.Values));
                    }
                }

                conn.Open();
                dbCommand.Connection = conn;
                dbCommand.CommandTimeout = CommandTimeOut == 0 ? DefaultTimeOut : CommandTimeOut;
                dbCommand.CommandType = commandType;
                dbCommand.CommandText = commandText;               
                dbCommand.ExecuteNonQuery();
                conn.Close();              
            }           
        }

        public static void ExecuteNonQuery(OleDbTransaction oleDbTransaction, CommandType commandType, string commandText, Parameters parameters = null)
        {
            OleDbCommand command = new OleDbCommand();

            // Add parameter
            if (parameters != null)
            {
                foreach (var par in parameters.GetParameters)
                {
                    command.Parameters.Add(new OleDbParameter(par.Name, par.Values));
                }
            }

            command.Transaction = oleDbTransaction;
            command.Connection = oleDbTransaction.Connection;
            command.CommandTimeout = CommandTimeOut == 0 ? DefaultTimeOut : CommandTimeOut;
            command.CommandType = commandType;
            command.CommandText = commandText;
            command.ExecuteNonQuery();
        }      

        public static DataSet ExecuteDataset(string connectionString, CommandType commandType, string commandText, Parameters parameters = null)
        {
            DataSet dataSet = new DataSet();
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                OleDbCommand command = new OleDbCommand();

                // Add parameter
                if (parameters != null)
                {
                    foreach (var par in parameters.GetParameters)
                    {
                        command.Parameters.Add(new SqlParameter(par.Name, par.Values));
                    }
                }

                command.Connection = conn;
                command.CommandType = commandType;
                command.CommandText = commandText;

                conn.Open();
                dataSet = ExecuteDataSet(command);
                conn.Close();
            }
            return dataSet;
        }

        private static DataSet ExecuteDataSet(OleDbCommand sqlCommand)
        {
            var ds = new DataSet();
            using (var dataAdapter = new OleDbDataAdapter(sqlCommand))
            {
                dataAdapter.Fill(ds);
            }
            return ds;
        }
    }

    public class Parameters
    {
        private IList<Parameter> _parameters;

        public Parameters()
        {
            _parameters = new List<Parameter>();
        }

        public void Add(Parameter paramter)
        {
            _parameters.Add(paramter);
        }

        public int Count { get { return _parameters.Count; } }

        public IList<Parameter> GetParameters { get { return _parameters; } }
    }

    public class Parameter
    {
        public string Name { get; set; }
        public object Values { get; set; }

    }
}
