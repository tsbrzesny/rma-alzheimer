using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RMA_Roots
{

    /* --------------------------------------
            DB rma_bankio
    */

    public class BIO_Bank
    {
        public int id { get; set; }
        public string shortname { get; set; }
        public string longname { get; set; }
        public string country { get; set; }
        public int? usescriptofid { get; set; }

        public static List<BIO_Bank> GetAll()
        {
            return DBAccess.dbBankIO.LoadDBItems<BIO_Bank>(@"select * from bio_bank");
        }

        public static BIO_Bank GetFromID(int id)
        {
            var temp = DBAccess.dbBankIO.LoadDBItems<BIO_Bank>(string.Format(@"select * from bio_bank where id = {0}", id.ToString()));
            return temp.FirstOrDefault();
        }

    }


    /* --------------------------------------
            DB misc
    */

    public class Misc_AppValue
    {
        public int id { get; set; }
        public string app_id { get; set; }
        public string key { get; set; }
        public string value { get; set; }


        public static List<Misc_AppValue> GetAll(string app_id, string key)
        {
            var sql = string.Format(@"select * from app_values where app_id = {0} and key = {1} order by id", DBAccess.PSQL_String(app_id), DBAccess.PSQL_String(key));
            return DBAccess.dbMisc.LoadDBItems<Misc_AppValue>(sql);
        }

        public static Misc_AppValue GetLast(string app_id, string key)
        {
            return GetAll(app_id, key).LastOrDefault();
        }

        public static int CreateNew(string app_id, string key, string value)
        {
            var sql = string.Format(@"insert into app_values (id, app_id, key, value) 
                                      values (default, {0}, {1}, {2}) returning id", 
                                    DBAccess.PSQL_String(app_id), DBAccess.PSQL_String(key), DBAccess.PSQL_String(value));
            return (int)DBAccess.dbMisc.SQL2O(sql);
        }

        public bool UpdateValue(string newValue)
        {
            try
            {
                var sql = string.Format("update app_values set value = {0} where id = {1}", DBAccess.PSQL_String(newValue), id);
                return (1 == DBAccess.dbMisc.SQLExec(sql));
            }
            catch (Exception)
            {
                return false;
            }
        }
    }


    public class Misc_AppException
    {
        public int id { get; set; }
        public DateTime created { get; set; }
        public string app_id { get; set; }
        public string key { get; set; }
        public string details { get; set; }
        public bool seen { get; set; }

        public static int CreateNew(string app_id, string key, string details)
        {
            var sql = string.Format(@"insert into app_exceptions (id, created, app_id, key, details, seen) 
                                      values (default, default, {0}, {1}, {2}, default) returning id",
                                    DBAccess.PSQL_String(app_id), DBAccess.PSQL_String(key), DBAccess.PSQL_String(details));
            return (int)DBAccess.dbMisc.SQL2O(sql);
        }
    }


    /* --------------------------------------
            DB stammdaten
    */

    public class Zahlungskonto
    {
        public int id { get; set; }
        public int mandant_ref { get; set; }
        public string bezeichnung { get; set; }
        public string bh_konto { get; set; }
        public string konto_nr { get; set; }
        public string bank_kurzname { get; set; }
        public string bank_blz { get; set; }
        public string bank_land { get; set; }
        public string currency { get; set; }
        public string esr_tn { get; set; }
        public string esr_kid { get; set; }
        public string konto_iban { get; set; }
        public string bic { get; set; }
        public string account_type { get; set; }
        public bool esr_enabled { get; set; }
        public bool is_standard { get; set; }

        // old links to bio_account, now replaced by reverse link bio_account.zahlungskonto_id & bio_account.is_active
        // how ever, these fields are still written by tha Booka app, because other parts rely on them (portal stuff)
        //
        // konto_nr_mt940 character varying(30),
        // mt940_from date,
        // mt940_to date,

        public static void UpdateMT940Section(int id, string joinedAccountName, bool? isActive)
        {
            if (isActive.HasValue && !isActive.Value || joinedAccountName == null)
            {
                DBAccess.dbStamm.SQLExec(string.Format(@"update zahlungskonto set konto_nr_mt940 = null, mt940_from = '20120223', mt940_to = '20120224' where id = {0}",
                                         DBAccess.PSQL_Int(id)));
            }

            else
            {
                DBAccess.dbStamm.SQLExec(string.Format(@"update zahlungskonto set konto_nr_mt940 = {1}, mt940_from = '20120223', mt940_to = null where id = {0}",
                                         DBAccess.PSQL_Int(id), DBAccess.PSQL_String(joinedAccountName)));
            }
        }

        // not (yet) needed:
        // dta_id character varying(20),
        // dta_abs character varying(20),
        // request_auszug_months integer NOT NULL DEFAULT 0,
        // export_type character varying(8),
        // valid_for_payments boolean,
        // associated_esr character varying(512),
        // gutschrift_app boolean DEFAULT false,
        // gutschrift_app_curr character varying(64),
        // gutschrift_app_type character varying(2) DEFAULT 'ar'::character varying,
        // letzte_buchung date,
        // letzter_mt940_abgleich date,
        // gutschrift_app_sepa_only boolean DEFAULT true,
        // remindersentdate timestamp without time zone,
    }

}
