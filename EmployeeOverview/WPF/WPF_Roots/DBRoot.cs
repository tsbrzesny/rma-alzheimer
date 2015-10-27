using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;

namespace WPF_Roots
{

    public class DBException : Exception
    {
        public DBException(string msg) : base(msg) { }
    }

    public abstract class ConnectionPool 
    {
        public abstract DbConnection GetConnection();
    };

    public class ConnectionPool<ConnectionType> : ConnectionPool 
           where ConnectionType: DbConnection, new()
    {
        private string name;
        private string connectionStr;
        private List<ConnectionType> pool = new List<ConnectionType>();
        private int initialThreadID;


        public ConnectionPool(string poolName, string connectionString)
        {
            this.name = poolName;
            this.connectionStr = connectionString;
            initialThreadID = Thread.CurrentThread.ManagedThreadId;
        }


        public override DbConnection GetConnection()
        {
            var isOnMainThread = (initialThreadID == Thread.CurrentThread.ManagedThreadId);
            ConnectionType conn = null;

            if (isOnMainThread)
                conn = pool.Find(_c => (_c.State == System.Data.ConnectionState.Closed || _c.State == System.Data.ConnectionState.Broken));

            if (conn == null)
            {
                conn = new ConnectionType();
                conn.ConnectionString = connectionStr;

                if (isOnMainThread)
                    pool.Add(conn);
            }

            // try to open the connection
            try
            {
                conn.Open();
            }
            catch (Exception ex)
            {
                var msg = string.Format("ConnectionPool.GetConnection():\nFehler beim Öffnen der DB-Verbindung (pool='{0}'):\n{1}", name, ex.Message);
                throw new DBException(msg);
            }

            //
            return conn;
        }


        public DbCommand GetCommand(DbConnection dbc, string sql)
        {
            var cmd = dbc.CreateCommand();
            cmd.CommandText = sql;
            return cmd;
        }



        /// SQLExec
        /// command execution - number of affected rows is returned
        /// 
        public int SQLExec(string sql, DbConnection dbc = null)
        {
            var useOwnConnection = (dbc == null);
            dbc = dbc ?? GetConnection();

            var cmd = GetCommand(dbc, sql);
            int affectedRows = 0;

            try
            {
                affectedRows = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                var msg = string.Format("ConnectionPool.SQLExec():\nFehler bei SQL-Abfrage (pool='{0}'):\n{1}\nSQL: {2}", name, ex.Message, sql);
                throw new DBException(msg);
            }
            finally
            {
                if (useOwnConnection)
                    dbc.Close();
            }

            return affectedRows;
        }

        /// SQL2O
        /// executes the given sql command & returns exactly 1 value.
        /// for example: 'insert into .. values (..) returning id'
        /// 
        public object SQL2O(string sql, DbConnection dbc = null)
        {
            var useOwnConnection = (dbc == null);
            dbc = dbc ?? GetConnection();

            var cmd = GetCommand(dbc, sql);
            DbDataReader dr = null;

            object result = null;
            try
            {
                dr = cmd.ExecuteReader();
                if (dr.Read())
                    result = dr.GetValue(0);
            }
            catch (Exception ex)
            {
                var msg = string.Format("ConnectionPool.SQL2O():\nFehler bei SQL-Abfrage (pool='{0}'):\n{1}\nSQL: {2}", name, ex.Message, sql);
                throw new DBException(msg);
            }
            finally
            {
                if (dr != null && !dr.IsClosed)
                    dr.Close();
                if (useOwnConnection)
                    dbc.Close();
            }

            return result;
        }


        /// SQL2LO
        /// executes the given sql command & returns a list of objects.
        /// 
        public List<object> SQL2LO(string sql, DbConnection dbc = null)
        {
            var useOwnConnection = (dbc == null);
            dbc = dbc ?? GetConnection();

            var cmd = GetCommand(dbc, sql);
            DbDataReader dr = null;

            var result = new List<Object>();
            try
            {
                dr = cmd.ExecuteReader();
                while (dr.Read())
                    result.Add(dr.GetValue(0));
            }
            catch (Exception ex)
            {
                var msg = string.Format("ConnectionPool.SQL2LO():\nFehler bei SQL-Abfrage (pool='{0}'):\n{1}\nSQL: {2}", name, ex.Message, sql);
                throw new DBException(msg);
            }
            finally
            {
                if (dr != null && !dr.IsClosed)
                    dr.Close();
                if (useOwnConnection)
                    dbc.Close();
            }

            return result;
        }


        /// SQL2LAO
        /// executes the given sql command & returns a list of records, coded as object[].
        /// 
        public List<object[]> SQL2LAO(string sql, DbConnection dbc = null)
        {
            var useOwnConnection = (dbc == null);
            dbc = dbc ?? GetConnection();

            var cmd = GetCommand(dbc, sql);
            DbDataReader dr = null;

            var result = new List<Object[]>();
            try
            {
                dr = cmd.ExecuteReader();
                var items = dr.FieldCount;

                while (dr.Read())
                {
                    var rec = new object[items];
                    dr.GetValues(rec);
                    result.Add(rec);
                }
            }
            catch (Exception ex)
            {
                var msg = string.Format("ConnectionPool.SQL2LAO():\nFehler bei SQL-Abfrage (pool='{0}'):\n{1}\nSQL: {2}", name, ex.Message, sql);
                throw new DBException(msg);
            }
            finally
            {
                if (dr != null && !dr.IsClosed)
                    dr.Close();
                if (useOwnConnection)
                    dbc.Close();
            }

            return result;
        }


        /// SQL2RD
        /// executes the given sql command & returns a List of RawDataRecords.
        /// 
        public List<RawDataRecord> SQL2RD(string sql, DbConnection dbc = null)
        {
            var useOwnConnection = (dbc == null);
            dbc = dbc ?? GetConnection();

            var cmd = GetCommand(dbc, sql);
            DbDataReader dr = null;
            object[] rowValues;

            var result = new List<RawDataRecord>();
            try
            {
                dr = cmd.ExecuteReader();

                // get column names
                var columnCount = dr.FieldCount;
                var columnNames = new string[columnCount];
                for (var i = 0; i < columnCount; i++)
                    columnNames[i] = dr.GetName(i);

                // get data rows
                rowValues = new object[columnCount];
                while (dr.Read())
                {
                    dr.GetValues(rowValues);
                    var rdr = new RawDataRecord(columnNames, rowValues);
                    result.Add(rdr);
                }
            }
            catch (Exception ex)
            {
                var msg = string.Format("ConnectionPool.SQL2RD():\nFehler bei SQL-Abfrage (pool='{0}'):\n{1}\nSQL: {2}", name, ex.Message, sql);
                throw new DBException(msg);
            }
            finally
            {
                if (dr != null && !dr.IsClosed)
                    dr.Close();
                if (useOwnConnection)
                    dbc.Close();
            }

            return result;
        }


        /// LoadDBItems
        /// loads a list of an arbitrary type T using reflection - 
        /// the fields of T must be properties & must have the same name as the fields in the DB table
        /// 
        public List<T> LoadDBItems<T>(String sql) where T : new()
        {
            var result = new List<T>();

            var conn = GetConnection();
            var cmd = GetCommand(conn, sql);
            DbDataReader dr = null;

            var recordCounter = 0;
            try
            {
                dr = cmd.ExecuteReader();
                int fieldCount = dr.FieldCount;

                // load column names
                var columnNames = new List<String>();
                for (int i = 0; i < fieldCount; i++)
                    columnNames.Add(dr.GetName(i));

                var rec = new Object[fieldCount];
                while (dr.Read())
                {
                    recordCounter++;
                    dr.GetValues(rec);
                    var jo = new JObject();
                    for (int i = 0; i < fieldCount; i++)
                    {
                        var value = rec[i];
                        if (value != null && value != DBNull.Value)
                            jo.Add(columnNames[i], new JValue(value));
                    }

                    var newT = new T();
                    JsonConvert.PopulateObject(jo.ToString(), newT);
                    result.Add(newT);
                }
            }
            catch (Exception ex)
            {
                var msg = string.Format("ConnectionPool.LoadDBItems():\nFehler bei SQL-Abfrage (pool='{0}'):\n{1}\nSQL: {2}", name, ex.Message, sql);
                throw new DBException(msg);
            }
            finally
            {
                if (dr != null && !dr.IsClosed)
                    dr.Close();
                conn.Close();
            }

            return result;
        }

    }


    public class RawDataRecord : Dictionary<string, object>
    {

        public RawDataRecord() : base() { }

        public RawDataRecord(string[] columnNames, object[] rowValues) : base()
        {
            var fieldCount = columnNames.Length;
            if (fieldCount != rowValues.Length)
                throw new DBException("DBRoot.RawDataRecord: Initialisations-Fehler.");

            for (var i = 0; i < fieldCount; i++)
                this[columnNames[i]] = rowValues[i];
        }

        public override string ToString()
        {
            var orderedParts = new List<string>();
            foreach (var columnName in this.Keys.OrderBy(k => k))
                orderedParts.Add(string.Format("{0}:{1}", columnName, this[columnName]));

            return string.Join("\n", orderedParts.ToArray());
        }


        // data getters

        public string AsString(string key, bool trim = true, bool nullAsEmpty = false)
        {
            if (!ContainsKey(key))
                throw new DBException(string.Format("DBRoot.RawDataRecord: Unbekannter Key {0}.", key == null ? "<NULL>" : "'" + key + "'"));

            var value = this[key];
            if (value == null || DBNull.Value.Equals(value))
                return nullAsEmpty ? String.Empty : null;

            if (trim)
                return value.ToString().Trim();

            return value.ToString();
        }


        public char AsChar(string key, char valueIfNull = '?')
        {
            if (!ContainsKey(key))
                throw new DBException(string.Format("DBRoot.RawDataRecord: Unbekannter Key {0}.", key == null ? "<NULL>" : "'" + key + "'"));

            var value = this[key];
            if (value is char)
                return (char)value;
            if (value == null || DBNull.Value.Equals(value))
                return valueIfNull;

            var result = value.ToString();
            if (result.Length > 0)
                return result[0];

            return valueIfNull;
        }


        public bool AsBoolean(string key, bool valueIfNull = false)
        {
            if (!ContainsKey(key))
                throw new DBException(string.Format("DBRoot.RawDataRecord: Unbekannter Key {0}.", key == null ? "<NULL>" : "'" + key + "'"));

            var value = this[key];
            if (value is Boolean)
                return (bool)value;

            else if (value == null || DBNull.Value.Equals(value))
                return valueIfNull;

            bool result = false;
            if (bool.TryParse(value.ToString(), out result))
                return result;

            return false;
        }

        public bool? AsBooleanNIL(string key)
        {
            if (!ContainsKey(key))
                throw new DBException(string.Format("DBRoot.RawDataRecord: Unbekannter Key {0}.", key == null ? "<NULL>" : "'" + key + "'"));

            var value = this[key];
            if (value is Boolean)
                return (bool)value;

            else if (value == null || DBNull.Value.Equals(value))
                return null;

            bool result = false;
            if (bool.TryParse(value.ToString(), out result))
                return result;

            return false;
        }


        public int AsInteger(string key, bool nullAsMinValue = false)
        {
            if (!ContainsKey(key))
                throw new DBException(string.Format("DBRoot.RawDataRecord: Unbekannter Key {0}.", key == null ? "<NULL>" : "'" + key + "'"));

            var value = this[key];
            if (value == null || DBNull.Value.Equals(value))
                return nullAsMinValue ? int.MinValue : 0;

            int result = 0;
            if (int.TryParse(value.ToString(), out result))
                return result;

            throw new DBException(string.Format("DBRoot.RawDataRecord.AsInteger: '{0}' kann nicht als Integer dargestellt werden (key = '{1}').", value.ToString(), key));
        }

        public int? AsIntegerNIL(string key)
        {
            if (!ContainsKey(key))
                throw new DBException(string.Format("DBRoot.RawDataRecord: Unbekannter Key {0}.", key == null ? "<NULL>" : "'" + key + "'"));

            var value = this[key];
            if (value == null || DBNull.Value.Equals(value))
                return null;

            int result = 0;
            if (int.TryParse(value.ToString(), out result))
                return result;

            throw new DBException(string.Format("DBRoot.RawDataRecord.AsIntegerNIL: '{0}' kann nicht als Integer dargestellt werden (key = '{1}').", value.ToString(), key));
        }


        public double AsDouble(string key)
        {
            if (!ContainsKey(key))
                throw new DBException(string.Format("DBRoot.RawDataRecord: Unbekannter Key {0}.", key == null ? "<NULL>" : "'" + key + "'"));

            var value = this[key];
            if (value == null || DBNull.Value.Equals(value))
                return 0.0;

            double result = 0.0;
            if (double.TryParse(value.ToString(), out result))
                return result;

            throw new DBException(string.Format("DBRoot.RawDataRecord.AsDouble: '{0}' kann nicht als Double dargestellt werden (key = '{1}').", value.ToString(), key));
        }


        public double? AsDoubleNIL(string key)
        {
            if (!ContainsKey(key))
                throw new DBException(string.Format("DBRoot.RawDataRecord: Unbekannter Key {0}.", key == null ? "<NULL>" : "'" + key + "'"));

            var value = this[key];
            if (value == null || DBNull.Value.Equals(value))
                return null;

            double result = 0.0;
            if (double.TryParse(value.ToString(), out result))
                return result;

            return null;
        }


        public DateTime AsDate(string key)
        {
            if (!ContainsKey(key))
                throw new DBException(string.Format("DBRoot.RawDataRecord: Unbekannter Key {0}.", key == null ? "<NULL>" : "'" + key + "'"));

            var value = this[key];
            DateTime result = DateTime.MinValue;

            if (value == null || DBNull.Value.Equals(value))
                return result;

            if (DateTime.TryParse(value.ToString(), out result))
                return result;

            throw new DBException(string.Format("DBRoot.RawDataRecord.AsDate: '{0}' kann nicht als DateTime dargestellt werden (key = '{1}').", value.ToString(), key));
        }


        public DateTime? AsDateNIL(string key)
        {
            if (!ContainsKey(key))
                throw new DBException(string.Format("DBRoot.RawDataRecord: Unbekannter Key {0}.", key == null ? "<NULL>" : "'" + key + "'"));

            var value = this[key];
            if (value == null || DBNull.Value.Equals(value))
                return null;

            DateTime result = DateTime.MinValue;
            if (DateTime.TryParse(value.ToString(), out result))
                return result;

            return null;
        }

    }

}
