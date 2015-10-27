using Npgsql;
using System;
using System.IO;
using System.Text;
using System.Xml.Linq;
using WPF_Roots;

namespace RMA_Roots
{
    public class DBAccess
    {
        public static ConnectionPool<NpgsqlConnection> dbBankIO = null;
        public static ConnectionPool<NpgsqlConnection> dbPayment = null;
        public static ConnectionPool<NpgsqlConnection> dbStamm = null;
        public static ConnectionPool<NpgsqlConnection> dbBilling = null;
        public static ConnectionPool<NpgsqlConnection> dbDoc = null;
        public static ConnectionPool<NpgsqlConnection> dbDocca = null;
        public static ConnectionPool<NpgsqlConnection> dbMisc = null;
        public static ConnectionPool<NpgsqlConnection> dbServicePortal = null;


        static DBAccess()
        {
            // initialize db connection pools
            dbBankIO = new ConnectionPool<NpgsqlConnection>("BankIO", AppConfig.GetItem("BankInDBString"));
            dbPayment = new ConnectionPool<NpgsqlConnection>("Payment", AppConfig.GetItem("PaymentDBString"));
            dbStamm = new ConnectionPool<NpgsqlConnection>("Stammdaten", AppConfig.GetItem("StammdatenDBString"));
            dbBilling = new ConnectionPool<NpgsqlConnection>("Billing", AppConfig.GetItem("Billing2DBString"));
            dbDoc = new ConnectionPool<NpgsqlConnection>("Doc", AppConfig.GetItem("DocDBString"));
            dbDocca = new ConnectionPool<NpgsqlConnection>("Docca", AppConfig.GetItem("DoccaDBString"));
            dbMisc = new ConnectionPool<NpgsqlConnection>("Misc", AppConfig.GetItem("MiscDBString"));
            dbServicePortal = new ConnectionPool<NpgsqlConnection>("ServicePortal", AppConfig.GetItem("ServicePortalDBString"));
        }

        public static ConnectionPool<NpgsqlConnection> GetLedgerPool(string mandantNameId)
        {
            var conStr = string.Format(AppConfig.GetItem("LedgerDBString"), mandantNameId);
            return new ConnectionPool<NpgsqlConnection>("SQLLedger", conStr);
        }

        // PostgreSQL data formatters


        public static string PSQL_Bool(bool value)
        {
            if (value)
                return "'true'";
            else
                return "'false'";
        }


        public static string PSQL_String(string str, bool nullAsNull = true, bool emptyAsNull = true)
        {
            //  prepares an input string to be used as a PostgreSQL string parameter
            //  - Nothing --> empty string / NULL
            //  - replace ' with ''
            //  - enclose in single '

            if (str == null)
            {
                if (nullAsNull)
                    return "NULL";
                str = "";
            }
            else
            {
                str.Trim();
                if (str.Length == 0 && emptyAsNull)
                    return "NULL";
            }

            str = str.Replace("'", "''");
            str = str.Replace(@"\", @"\\");
            return "E'" + str + "'";
        }


        public static string PSQL_XML(XDocument xDoc)
        {
            if (xDoc == null)
                return "NULL";

            var sb = new StringBuilder();
            using (var sw = new StringWriter(sb))
            {
                xDoc.Save(sw, SaveOptions.None);
            }
            return PSQL_String(sb.ToString().Replace("encoding=\"utf-16\"", "encoding=\"utf-8\""));
        }


        public static string PSQL_Date(DateTime d)
        {
            if (d == null || d == DateTime.MinValue)
                return "NULL";

            return PSQL_String(d.ToString("yyyy-MM-dd"));
        }


        public static string PSQL_DateTime(DateTime d)
        {
            if (d == null || d == DateTime.MinValue)
                return "NULL";

            return PSQL_String(d.ToString("yyyy-MM-dd HH:mm:ss"));
        }


        public static string PSQL_Int(int i, bool minValueAsNull = true)
        {
            if (i == int.MinValue && minValueAsNull)
                return "NULL";

            return i.ToString();
        }

        public static string PSQL_IntNull(int? i)
        {
            if (i.HasValue)
                return i.ToString();
            return "NULL";
        }

        public static string PSQL_EqualsIntNull(int? i)
        {
            if (i.HasValue)
                return "= " + i.ToString();
            return "is NULL";
        }

    }

}
