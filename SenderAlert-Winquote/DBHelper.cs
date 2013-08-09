
using System;
using System.Data;
using System.Xml;
using System.Data.OleDb;
using System.Collections;

namespace SenderSMS_TS8
{
    public sealed class DBHelper
    {
        private static DBHelper me;
        private static OleDbConnection cn;
        private DBHelper()
        {

        }

		public static void Close()
		{
			if (cn != null)
			{
				cn.Close();
				cn = null;
			}
			if (me != null)
			{
				me =null;
			}
		}

        public static DBHelper getInstance(string connectionString)
        {
            if (cn == null)
            {
                cn = new OleDbConnection(connectionString);
            }
            if (cn.State != ConnectionState.Open)
            {
                cn.Open();
            }
            if (me == null)
            {
                me = new DBHelper();
            }
            return me;
        }

        public DataSet ExecuteDataset(string commandText)
        {
            //create a command and prepare it for execution
            OleDbCommand cmd = new OleDbCommand();
            PrepareCommand(cmd, cn, CommandType.Text, commandText);

            //create the DataAdapter & DataSet
            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            DataSet ds = new DataSet();

            //fill the DataSet using default values for DataTable names, etc.
            da.Fill(ds);

            //return the dataset
            return ds;
        }


        public int ExecuteNonQuery(string commandText)
        {
            //create a command and prepare it for execution
            OleDbCommand cmd = new OleDbCommand();
            PrepareCommand(cmd, cn, CommandType.Text, commandText);

            //finally, execute the command.
            return cmd.ExecuteNonQuery();
        }

        public object ExecuteScalar(string commandText)
        {
            //create a command and prepare it for execution
            OleDbCommand cmd = new OleDbCommand();
            PrepareCommand(cmd, cn, CommandType.Text, commandText);

            //execute the command & return the results
            return cmd.ExecuteScalar();
            
        }

          private  void PrepareCommand(OleDbCommand command, OleDbConnection connection, CommandType commandType, string commandText)
         {
           
            //associate the connection with the command
            command.Connection = connection;

            //set the command text (stored procedure name or OleDb statement)
            command.CommandText = commandText;
                       
            //set the command type
            command.CommandType = commandType;
                       
            return;
        }
    }

}
