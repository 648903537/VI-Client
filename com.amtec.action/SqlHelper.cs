using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Collections;
using System.Data.OleDb;
using System.Reflection;
using System.IO;

namespace com.amtec.action
{
    /// <summary>
    /// The SqlHelper class is intended to encapsulate high performance, 
    /// scalable best practices for common uses of SqlClient.
    /// </summary>
    public class SqlHelper
    {
        private string dbPath = "";
        public SqlHelper(string strDBPath)
        {
            dbPath = strDBPath;
        }
        private string connectionString
        {
            get
            {
                //string dbPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName) + @"\OATest.xlsx";
                //return "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" + dbPath + ";Extended Properties='Excel 8.0'";
                return "Provider=Microsoft.Ace.OleDb.12.0;" + "data source=" + dbPath + ";Extended Properties='Excel 12.0';";//HDR=NO; IMEX=1
            }
        }

        private int commandTimeout = 60;

        private Hashtable parmCache = Hashtable.Synchronized(new Hashtable());

        /// <summary>
        /// Execute a OleDbCommand (that returns no resultset) against the database specified in the connection string 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new OleDbParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connectionString">a valid connection string for a OleDbConnection</param>
        /// <param name="commandType">the CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">the stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">an array of SqlParamters used to execute the command</param>
        /// <returns>an int representing the number of rows affected by the command</returns>
        public int ExecuteNonQuery(string cmdText, params OleDbParameter[] commandParameters)
        {

            OleDbCommand cmd = new OleDbCommand();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                try
                {
                    PrepareCommand(cmd, conn, null, cmdText, commandParameters);
                    int val = cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                    return val;
                }
                catch (Exception ex)
                {
                    conn.Close();
                    throw new Exception(ex.Message);
                }
            }
        }

        /// <summary>
        /// Execute a OleDbCommand (that returns no resultset) against an existing database connection 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new OleDbParameter("@prodid", 24));
        /// </remarks>
        /// <param name="conn">an existing database connection</param>
        /// <param name="commandType">the CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">the stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">an array of SqlParamters used to execute the command</param>
        /// <returns>an int representing the number of rows affected by the command</returns>
        public int ExecuteNonQuery(OleDbConnection connection, string cmdText, params OleDbParameter[] commandParameters)
        {

            OleDbCommand cmd = new OleDbCommand();

            PrepareCommand(cmd, connection, null, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        /// <summary>
        /// Execute a OleDbCommand (that returns no resultset) using an existing SQL Transaction 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  int result = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new OleDbParameter("@prodid", 24));
        /// </remarks>
        /// <param name="trans">an existing sql transaction</param>
        /// <param name="commandType">the CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">the stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">an array of SqlParamters used to execute the command</param>
        /// <returns>an int representing the number of rows affected by the command</returns>
        public int ExecuteNonQuery(OleDbTransaction trans, string cmdText, params OleDbParameter[] commandParameters)
        {
            OleDbCommand cmd = new OleDbCommand();
            PrepareCommand(cmd, trans.Connection, trans, cmdText, commandParameters);
            int val = cmd.ExecuteNonQuery();
            cmd.Parameters.Clear();
            return val;
        }

        /// <summary>
        /// 填充DataSet
        /// </summary>
        /// <param name="connectionString"></param>
        /// <param name="CommandType.Text"></param>
        /// <param name="cmdText"></param>
        /// <param name="commandParameters"></param>
        /// <returns></returns>
        public DataSet FillDataReaderToDataSet(string cmdText, params OleDbParameter[] commandParameters)
        {
            OleDbDataReader sdr = ExecuteReader(cmdText, commandParameters);
            DataSet ds = DataReaderToDataSet(sdr);
            return ds;
        }

        #region FillDataset
        /// <summary>
        /// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the database specified in 
        /// the connection string. 
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  FillDataset(connString, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"});
        /// </remarks>
        /// <param name="connectionString">A valid connection string for a OleDbConnection</param>
        /// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">The stored procedure name or T-SQL command</param>
        /// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
        /// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
        /// by a user defined name (probably the actual table name)</param>
        public void FillDataset(string connectionString, CommandType commandType, string commandText, DataSet dataSet, string[] tableNames)
        {
            if (connectionString == null || connectionString.Length == 0) throw new ArgumentNullException("connectionString");
            if (dataSet == null) throw new ArgumentNullException("dataSet");

            // Create & open a OleDbConnection, and dispose of it after we are done
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Call the overload that takes a connection in place of the connection string
                FillDataset(connection, commandType, commandText, dataSet, tableNames);
            }
        }

        /// <summary>
        /// Execute a OleDbCommand (that returns a resultset) against the database specified in the connection string 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  FillDataset(connString, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"}, new OleDbParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connectionString">A valid connection string for a OleDbConnection</param>
        /// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">The stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">An array of SqlParamters used to execute the command</param>
        /// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
        /// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
        /// by a user defined name (probably the actual table name)
        /// </param>
        public void FillDataset(string connectionString, CommandType commandType,
            string commandText, DataSet dataSet, string[] tableNames,
            params OleDbParameter[] commandParameters)
        {
            if (connectionString == null || connectionString.Length == 0) throw new ArgumentNullException("connectionString");
            if (dataSet == null) throw new ArgumentNullException("dataSet");
            // Create & open a OleDbConnection, and dispose of it after we are done
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                // Call the overload that takes a connection in place of the connection string
                FillDataset(connection, commandType, commandText, dataSet, tableNames, commandParameters);
            }
        }

        /// <summary>
        /// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the provided OleDbConnection. 
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  FillDataset(conn, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"});
        /// </remarks>
        /// <param name="connection">A valid OleDbConnection</param>
        /// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">The stored procedure name or T-SQL command</param>
        /// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
        /// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
        /// by a user defined name (probably the actual table name)
        /// </param>    
        public void FillDataset(OleDbConnection connection, CommandType commandType,
            string commandText, DataSet dataSet, string[] tableNames)
        {
            FillDataset(connection, commandType, commandText, dataSet, tableNames, null);
        }

        /// <summary>
        /// Execute a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  FillDataset(conn, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"}, new OleDbParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connection">A valid OleDbConnection</param>
        /// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">The stored procedure name or T-SQL command</param>
        /// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
        /// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
        /// by a user defined name (probably the actual table name)
        /// </param>
        /// <param name="commandParameters">An array of SqlParamters used to execute the command</param>
        public void FillDataset(OleDbConnection connection, CommandType commandType,
            string commandText, DataSet dataSet, string[] tableNames,
            params OleDbParameter[] commandParameters)
        {
            FillDataset(connection, null, commandType, commandText, dataSet, tableNames, commandParameters);
        }

        /// <summary>
        /// Execute a OleDbCommand (that returns a resultset and takes no parameters) against the provided OleDbTransaction. 
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  FillDataset(trans, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"});
        /// </remarks>
        /// <param name="transaction">A valid OleDbTransaction</param>
        /// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">The stored procedure name or T-SQL command</param>
        /// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
        /// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
        /// by a user defined name (probably the actual table name)
        /// </param>
        public void FillDataset(OleDbTransaction transaction, CommandType commandType,
            string commandText,
            DataSet dataSet, string[] tableNames)
        {
            FillDataset(transaction, commandType, commandText, dataSet, tableNames, null);
        }

        /// <summary>
        /// Execute a OleDbCommand (that returns a resultset) against the specified OleDbTransaction
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  FillDataset(trans, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"}, new OleDbParameter("@prodid", 24));
        /// </remarks>
        /// <param name="transaction">A valid OleDbTransaction</param>
        /// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">The stored procedure name or T-SQL command</param>
        /// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
        /// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
        /// by a user defined name (probably the actual table name)
        /// </param>
        /// <param name="commandParameters">An array of SqlParamters used to execute the command</param>
        public void FillDataset(OleDbTransaction transaction, CommandType commandType,
            string commandText, DataSet dataSet, string[] tableNames,
            params OleDbParameter[] commandParameters)
        {
            FillDataset(transaction.Connection, transaction, commandType, commandText, dataSet, tableNames, commandParameters);
        }


        /// <summary>
        /// Private helper method that execute a OleDbCommand (that returns a resultset) against the specified OleDbTransaction and OleDbConnection
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  FillDataset(conn, trans, CommandType.StoredProcedure, "GetOrders", ds, new string[] {"orders"}, new OleDbParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connection">A valid OleDbConnection</param>
        /// <param name="transaction">A valid OleDbTransaction</param>
        /// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">The stored procedure name or T-SQL command</param>
        /// <param name="dataSet">A dataset wich will contain the resultset generated by the command</param>
        /// <param name="tableNames">This array will be used to create table mappings allowing the DataTables to be referenced
        /// by a user defined name (probably the actual table name)
        /// </param>
        /// <param name="commandParameters">An array of SqlParamters used to execute the command</param>
        private void FillDataset(OleDbConnection connection, OleDbTransaction transaction, CommandType commandType,
            string commandText, DataSet dataSet, string[] tableNames,
            params OleDbParameter[] commandParameters)
        {
            if (connection == null) throw new ArgumentNullException("connection");
            if (dataSet == null) throw new ArgumentNullException("dataSet");

            // Create a command and prepare it for execution
            OleDbCommand command = new OleDbCommand();
            bool mustCloseConnection = false;
            PrepareCommand(command, connection, transaction, commandType, commandText, commandParameters, out mustCloseConnection);

            // Create the DataAdapter & DataSet
            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(command))
            {

                // Add the table mappings specified by the user
                if (tableNames != null && tableNames.Length > 0)
                {
                    string tableName = "Table";
                    for (int index = 0; index < tableNames.Length; index++)
                    {
                        if (tableNames[index] == null || tableNames[index].Length == 0) throw new ArgumentException("The tableNames parameter must contain a list of tables, a value was provided as null or empty string.", "tableNames");
                        dataAdapter.TableMappings.Add(tableName, tableNames[index]);
                        tableName = "Table" + (index + 1).ToString();
                    }
                }

                // Fill the DataSet using default values for DataTable names, etc
                dataAdapter.Fill(dataSet);

                // Detach the SqlParameters from the command object, so they can be used again
                command.Parameters.Clear();
            }

            if (mustCloseConnection)
                connection.Close();
        }
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        private DataSet DataReaderToDataSet(IDataReader reader)
        {
            DataTable table = new DataTable();
            int fieldCount = reader.FieldCount;
            for (int i = 0; i < fieldCount; i++)
            {
                table.Columns.Add(reader.GetName(i), reader.GetFieldType(i));
            }
            table.BeginLoadData();
            object[] values = new object[fieldCount];
            while (reader.Read())
            {
                reader.GetValues(values);
                table.LoadDataRow(values, true);
            }
            table.EndLoadData();
            DataSet ds = new DataSet();
            ds.Tables.Add(table);
            return ds;
        }


        /// <summary>
        /// Execute a OleDbCommand that returns a resultset against the database specified in the connection string 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  OleDbDataReader r = ExecuteReader(connString, CommandType.StoredProcedure, "PublishOrders", new OleDbParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connectionString">a valid connection string for a OleDbConnection</param>
        /// <param name="commandType">the CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">the stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">an array of SqlParamters used to execute the command</param>
        /// <returns>A OleDbDataReader containing the results</returns>
        public OleDbDataReader ExecuteReader(string cmdText, params OleDbParameter[] commandParameters)
        {
            OleDbCommand cmd = new OleDbCommand();
            OleDbConnection conn = new OleDbConnection(connectionString);

            // we use a try/catch here because if the method throws an exception we want to 
            // close the connection throw code, because no datareader will exist, hence the 
            // commandBehaviour.CloseConnection will not work
            try
            {
                PrepareCommand(cmd, conn, null, cmdText, commandParameters);
                OleDbDataReader rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                cmd.Parameters.Clear();
                return rdr;
            }
            catch (Exception ex)
            {
                conn.Close();
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Execute a OleDbCommand that returns the first column of the first record against the database specified in the connection string 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  Object obj = ExecuteScalar(connString, CommandType.StoredProcedure, "PublishOrders", new OleDbParameter("@prodid", 24));
        /// </remarks>
        /// <param name="connectionString">a valid connection string for a OleDbConnection</param>
        /// <param name="commandType">the CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">the stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">an array of SqlParamters used to execute the command</param>
        /// <returns>An object that should be converted to the expected type using Convert.To{Type}</returns>
        public object ExecuteScalar(string cmdText, params OleDbParameter[] commandParameters)
        {
            OleDbCommand cmd = new OleDbCommand();

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                PrepareCommand(cmd, connection, null, cmdText, commandParameters);
                object val = cmd.ExecuteScalar();
                cmd.Parameters.Clear();
                return val;
            }
        }

        /// <summary>
        /// Execute a OleDbCommand that returns the first column of the first record against an existing database connection 
        /// using the provided parameters.
        /// </summary>
        /// <remarks>
        /// e.g.:  
        ///  Object obj = ExecuteScalar(connString, CommandType.StoredProcedure, "PublishOrders", new OleDbParameter("@prodid", 24));
        /// </remarks>
        /// <param name="conn">an existing database connection</param>
        /// <param name="commandType">the CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">the stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">an array of SqlParamters used to execute the command</param>
        /// <returns>An object that should be converted to the expected type using Convert.To{Type}</returns>
        public object ExecuteScalar(OleDbConnection connection, string cmdText, params OleDbParameter[] commandParameters)
        {

            OleDbCommand cmd = new OleDbCommand();

            PrepareCommand(cmd, connection, null, cmdText, commandParameters);
            object val = cmd.ExecuteScalar();
            cmd.Parameters.Clear();
            return val;
        }

        /// <summary>
        /// add parameter array to the cache
        /// </summary>
        /// <param name="cacheKey">Key to the parameter cache</param>
        /// <param name="cmdParms">an array of SqlParamters to be cached</param>
        public void CacheParameters(string cacheKey, params OleDbParameter[] commandParameters)
        {
            parmCache[cacheKey] = commandParameters;
        }

        /// <summary>
        /// Retrieve cached parameters
        /// </summary>
        /// <param name="cacheKey">key used to lookup parameters</param>
        /// <returns>Cached SqlParamters array</returns>
        public OleDbParameter[] GetCachedParameters(string cacheKey)
        {
            OleDbParameter[] cachedParms = (OleDbParameter[])parmCache[cacheKey];

            if (cachedParms == null)
                return null;

            OleDbParameter[] clonedParms = new OleDbParameter[cachedParms.Length];

            for (int i = 0, j = cachedParms.Length; i < j; i++)
                clonedParms[i] = (OleDbParameter)((ICloneable)cachedParms[i]).Clone();

            return clonedParms;
        }

        /// <summary>
        /// This method opens (if necessary) and assigns a connection, transaction, command type and parameters 
        /// to the provided command
        /// </summary>
        /// <param name="command">The OleDbCommand to be prepared</param>
        /// <param name="connection">A valid OleDbConnection, on which to execute this command</param>
        /// <param name="transaction">A valid OleDbTransaction, or 'null'</param>
        /// <param name="commandType">The CommandType (stored procedure, text, etc.)</param>
        /// <param name="commandText">The stored procedure name or T-SQL command</param>
        /// <param name="commandParameters">An array of SqlParameters to be associated with the command or 'null' if no parameters are required</param>
        /// <param name="mustCloseConnection"><c>true</c> if the connection was opened by the method, otherwose is false.</param>
        private void PrepareCommand(OleDbCommand command, OleDbConnection connection, OleDbTransaction transaction, CommandType commandType, string commandText, OleDbParameter[] commandParameters, out bool mustCloseConnection)
        {
            if (command == null) throw new ArgumentNullException("command");
            if (commandText == null || commandText.Length == 0) throw new ArgumentNullException("commandText");
            command.CommandTimeout = commandTimeout;


            // If the provided connection is not open, we will open it
            if (connection.State != ConnectionState.Open)
            {
                mustCloseConnection = true;
                connection.Open();
            }
            else
            {
                mustCloseConnection = false;
            }

            // Associate the connection with the command
            command.Connection = connection;

            // Set the command text (stored procedure name or SQL statement)
            command.CommandText = commandText;

            // If we were provided a transaction, assign it
            if (transaction != null)
            {
                if (transaction.Connection == null) throw new ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction");
                command.Transaction = transaction;
            }

            // Set the command type
            command.CommandType = commandType;

            // Attach the command parameters if they are provided
            if (commandParameters != null)
            {
                AttachParameters(command, commandParameters);
            }
            return;
        }

        private void AttachParameters(OleDbCommand command, OleDbParameter[] commandParameters)
        {
            if (command == null) throw new ArgumentNullException("command");
            if (commandParameters != null)
            {
                foreach (OleDbParameter p in commandParameters)
                {
                    if (p != null)
                    {
                        // Check for derived output value with no value assigned
                        if ((p.Direction == ParameterDirection.InputOutput ||
                            p.Direction == ParameterDirection.Input) &&
                            (p.Value == null))
                        {
                            p.Value = DBNull.Value;
                        }
                        command.Parameters.Add(p);
                    }
                }
            }
        }

        private void PrepareCommand(OleDbCommand cmd, OleDbConnection conn, OleDbTransaction trans, string cmdText, OleDbParameter[] cmdParms)
        {
            if (conn.State != ConnectionState.Open)
                conn.Open();
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            cmd.CommandTimeout = commandTimeout;
            if (trans != null)
                cmd.Transaction = trans;
            cmd.CommandType = CommandType.Text;//CommandType.Text;
            if (cmdParms != null)
            {
                foreach (OleDbParameter parameter in cmdParms)
                {
                    if ((parameter.Direction == ParameterDirection.InputOutput || parameter.Direction == ParameterDirection.Input) &&
                        (parameter.Value == null))
                    {
                        parameter.Value = DBNull.Value;
                    }

                    cmd.Parameters.Add(parameter);

                }
            }
        }

        private void SelectPrepareCommand(OleDbCommand cmd, OleDbConnection conn, OleDbTransaction trans, string cmdText, OleDbParameter[] cmdParms)
        {
            if (conn.State != ConnectionState.Open)
                conn.Open();
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            cmd.CommandTimeout = commandTimeout;
            if (trans != null)
                cmd.Transaction = trans;
            cmd.CommandType = CommandType.Text;//CommandType.Text;
            if (cmdParms != null)
            {
                foreach (OleDbParameter parameter in cmdParms)
                {
                    if ((parameter.Direction == ParameterDirection.InputOutput || parameter.Direction == ParameterDirection.Input) &&
                        (parameter.Value == null))
                    {
                        parameter.Value = null;
                    }
                    else
                    {
                        cmd.Parameters.Add(parameter);
                    }
                }
            }
        }

        /// <summary>
        /// 仅支持SQL语句查询，返回DataSet
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="SQLString">查询语句</param>
        /// <returns>DataSet</returns>
        public DataSet Fill(string SQLString, params OleDbParameter[] cmdParms)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand cmd = new OleDbCommand();
                SelectPrepareCommand(cmd, connection, null, SQLString, cmdParms);
                using (OleDbDataAdapter da = new OleDbDataAdapter(cmd))
                {
                    DataSet ds = new DataSet();
                    try
                    {
                        da.Fill(ds, "ds");
                        cmd.Parameters.Clear();
                    }
                    catch (System.Data.SqlClient.SqlException ex)
                    {
                        throw new Exception(ex.Message);
                    }
                    finally
                    {
                        cmd.Dispose();
                        connection.Close();
                    }
                    return ds;
                }
            }
        }

        public void Fill(string SQLString, DataSet ds, string tableName, params OleDbParameter[] cmdParms)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand cmd = new OleDbCommand();
                SelectPrepareCommand(cmd, connection, null, SQLString, cmdParms);
                using (OleDbDataAdapter da = new OleDbDataAdapter(cmd))
                {
                    try
                    {
                        da.Fill(ds, tableName);
                        cmd.Parameters.Clear();
                    }
                    catch (System.Data.SqlClient.SqlException ex)
                    {
                        throw new Exception(ex.Message);
                    }
                    finally
                    {
                        cmd.Dispose();
                        connection.Close();
                    }
                }
            }
        }

        /// <summary>
        /// DataSet转换为泛型集合 (
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="p_DataSet"></param>
        /// <param name="p_TableIndex"></param>
        /// <returns></returns>
        public List<T> DataSetToList<T>(DataSet p_DataSet, int p_TableIndex)
        {
            IList<T> ilist = DataSetToIList<T>(p_DataSet, p_TableIndex);
            List<T> result = new List<T>(ilist);
            return result;
        }

        /// <summary> 
        /// DataSet转换为泛型集合 (Add by Jackie.Shi 2008-09-01修改)
        /// </summary> 
        /// <typeparam name="T"></typeparam> 
        /// <param name="p_DataSet">DataSet</param> 
        /// <param name="p_TableIndex">待转换数据表索引</param> 
        /// <returns></returns> 
        public IList<T> DataSetToIList<T>(DataSet p_DataSet, int p_TableIndex)
        {
            // 返回值初始化 
            IList<T> result = new List<T>();

            if (p_DataSet == null || p_DataSet.Tables.Count < 0)
                return result;

            if (p_TableIndex > p_DataSet.Tables.Count - 1)
                return result;

            if (p_TableIndex < 0)
                p_TableIndex = 0;

            DataTable p_Data = p_DataSet.Tables[p_TableIndex];

            for (int j = 0; j < p_Data.Rows.Count; j++)
            {
                T _t = (T)Activator.CreateInstance(typeof(T));
                PropertyInfo[] propertys = _t.GetType().GetProperties();
                foreach (PropertyInfo pi in propertys)
                {
                    for (int i = 0; i < p_Data.Columns.Count; i++)
                    {
                        // 属性与字段名称一致的进行赋值 
                        if (pi.Name.ToUpper().Equals(p_Data.Columns[i].ColumnName.ToUpper()))
                        {
                            // 数据库NULL值单独处理 
                            if (p_Data.Rows[j][i] != DBNull.Value && p_Data.Rows[j][i] != null)
                            {
                                if (pi.PropertyType == typeof(Boolean))
                                {
                                    pi.SetValue(_t, Convert.ToBoolean(p_Data.Rows[j][i]), null);
                                }
                                else if (pi.PropertyType == typeof(Int32))
                                {
                                    pi.SetValue(_t, Convert.ToInt32(p_Data.Rows[j][i]), null);
                                }
                                else if (pi.PropertyType == typeof(Decimal))
                                {
                                    pi.SetValue(_t, Convert.ToDecimal(p_Data.Rows[j][i]), null);
                                }
                                //Add By : Jet Xu
                                //Add Date:2008-11-24
                                //Add Desc:增加Int32?类型的处理
                                else if (pi.PropertyType == typeof(Int32?))
                                {
                                    pi.SetValue(_t, Convert.ToInt32(p_Data.Rows[j][i]), null);
                                }
                                //End Add
                                else
                                {
                                    pi.SetValue(_t, p_Data.Rows[j][i], null);
                                }
                            }
                            else
                                pi.SetValue(_t, null, null);
                            break;
                        }
                    }
                }
                result.Add(_t);
            }
            return result;
        }

        /// <summary> 
        /// DataSet转换为泛型集合 
        /// </summary> 
        /// <typeparam name="T"></typeparam> 
        /// <param name="p_DataSet">DataSet</param> 
        /// <param name="p_TableName">待转换数据表名称</param> 
        /// <returns></returns> 
        public IList<T> DataSetToIList<T>(DataSet p_DataSet, string p_TableName)
        {
            int _TableIndex = 0;

            if (p_DataSet == null || p_DataSet.Tables.Count < 0)
                return new List<T>();

            if (string.IsNullOrEmpty(p_TableName))
                return new List<T>();

            for (int i = 0; i < p_DataSet.Tables.Count; i++)
            {
                // 获取Table名称在Tables集合中的索引值 
                if (p_DataSet.Tables[i].TableName.ToUpper().Equals(p_TableName.ToUpper()))
                {
                    _TableIndex = i;
                    break;
                }
            }
            return DataSetToIList<T>(p_DataSet, _TableIndex);
        }
    }
}
