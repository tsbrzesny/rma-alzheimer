using Npgsql;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Documents;
using WPF_Roots;

namespace RMA_Roots
{
    public static class AppConfig
    {

        // configuration management - quite the same as in RMA2_Roots

        private static Dictionary<string, object> allItems = new Dictionary<string, object>();

        // this ini item, if present, specifies the path to the global ini file
        // Note the local items override global ones
        private const string Key_GlobalIniFile = "$GlobalINI";


        public static bool localIniFileFound = false;
        public static bool globalIniFileFound = false;
        public static bool dbRecordFound = false;
        //
        public static void LoadAppConfiguration(string appId = null)
        {
            // load local application configuration
            Dictionary<string, object> localIni = null;

            var entryAssembly = Assembly.GetEntryAssembly();
            if (ReadIniFile(Regex.Replace(entryAssembly.Location, @"\.exe$", @".ini"), ref localIni))
                localIniFileFound = true;

            // load global ini
            string globalIniPath = null;
            if (localIniFileFound && localIni.ContainsKey(Key_GlobalIniFile))
                globalIniPath = (string)localIni[Key_GlobalIniFile];
            globalIniFileFound = ReadIniFile(globalIniPath, ref allItems);

            // try to load app configuration from the misc table, if
            // - no local ini file was found
            // - the global ini file was successfully loaded (necessary to access the dbs)
            // - an appId is given (as key into the misc.app_value table)
            if (!localIniFileFound && globalIniFileFound && appId != null)
            {
                var configRecord = Misc_AppValue.GetLast(appId, "AppConfig");
                if (configRecord != null)
                {
                    dbRecordFound = true;
                    ParseIniContent(configRecord.value, ref localIni);
                }
            }

            // values from the local ini files override global ones
            if (localIniFileFound)
            {
                foreach (var _kv in localIni)
                    allItems[_kv.Key] = _kv.Value;
            }
        }


        private static bool ReadIniFile(string fullPath, ref Dictionary<string, object> intoRepository)
        {
            if (!File.Exists(fullPath))
                return false;

            var content = File.ReadAllText(fullPath);
            ParseIniContent(content, ref intoRepository);
            return true;
        }

        public static Section TestConfiguration(string configText, string[] allowedKeys = null)
        {
            try
            {
                Dictionary<string, object> testRepository = null;
                ParseIniContent(configText, ref testRepository, allowedKeys);

                // null means: no errors detected
                return null;
            }
            catch (Exception e)
            {
                if (formattedError != null)
                    return formattedError;
                throw e;
            }
        }

        private static Section formattedError = null;
        //
        private static void ParseIniContent(string content, ref Dictionary<string, object> intoRepository, string[] allowedKeys = null)
        {
            if (intoRepository == null)
                intoRepository = new Dictionary<string, object>();
            formattedError = null;

            var lineNumber = 0;
            foreach (var _line in Regex.Split(content, "\r|\n"))
            {
                lineNumber++;
                var line = _line.Trim();
                if (line.Length == 0 || line.StartsWith("#"))
                    continue;

                // the first '=' on the line splits into key/value
                var kv = line.Split(new char[] { '=' }, 2);
                if (kv.Count() != 2)
                {
                    var errorHeader = string.Format("Fehler auf Zeile {0}: Zeile ist kein Kommentar ('#') und enthält kein '=' (key = value).", lineNumber);
                    formattedError = (Section)RTB_Support.CreateDocElement("Section", 
                                        string.Format("<Paragraph FontWeight='Bold'>{0}</Paragraph>", RTB_Support.XAMLEsc(errorHeader)), "0", null);
                    throw new ApplicationException(errorHeader);
                }
                var key = kv[0].Trim();
                var value = kv[1].Trim();

                if (allowedKeys != null && !allowedKeys.Any(_ak => key == _ak || Regex.IsMatch(key, string.Format(@"^{0}\(.*\)$", Regex.Escape(_ak)))))
                {
                    var errorHeader = string.Format("Fehler auf Zeile {0}: '{1}' ist kein erlaubter Schlüssel.", lineNumber, key);
                    var allNoWrap = (from _ak in allowedKeys select string.Join("\u2060", _ak.ToCharArray())).ToArray();
                    var errorComment = string.Format("(erlaubte Schlüssel sind: {0})", string.Join(", ", allNoWrap));
                    formattedError = (Section)RTB_Support.CreateDocElement("Section",
                                        string.Format("<Paragraph FontWeight='Bold'>{0}</Paragraph><Paragraph Foreground='Silver'>{1}</Paragraph>",
                                        RTB_Support.XAMLEsc(errorHeader), RTB_Support.XAMLEsc(errorComment)), "0", null);
                    throw new ApplicationException(errorHeader);
                }

                if (key.StartsWith("@"))
                {
                    if (intoRepository.ContainsKey(key))
                        ((List<string>)intoRepository[key]).Add(value);
                    else
                        intoRepository[key] = new List<string>() { value };
                }

                else if (key.StartsWith("%"))
                {
                    var m = Regex.Match(key, @"(%\w+)\((\w+)\)");
                    if (m.Success)
                    {
                        key = m.Groups[1].Value;
                        var subkey = m.Groups[2].Value.Trim();

                        if (!intoRepository.ContainsKey(key))
                            intoRepository[key] = new Dictionary<string, string>();

                        ((Dictionary<string, string>)intoRepository[key])[subkey] = value;
                    }
                }

                else
                {
                    intoRepository["$" + key] = value;
                }
            }

        }


        public static string GetItem(string key, string defaultValue = null)
        {
            if (key == null)
                return defaultValue;
            else
                key = "$" + key;

            if (allItems.ContainsKey(key))
                return (string)allItems[key];
            return defaultValue;
        }

        public static void SetItem(string key, string value)
        {
            if (key == null)
                throw new ApplicationException("Key cannot be Null.");
            allItems["$" + key] = value;
        }

        public static List<string> GetList(string key, bool returnEmptyListIfKeyNotFound = true)
        {
            if (key == null)
                return null;
            else
                key = "@" + key;

            if (allItems.ContainsKey(key))
                return (List<string>)allItems[key];

            if (returnEmptyListIfKeyNotFound)
                return new List<string>();
            return null;
        }

        public static void SetList(string key, List<string> value)
        {
            if (key == null)
                throw new ApplicationException("Key cannot be Null.");
            allItems["@" + key] = value;
        }

        public static Dictionary<string, string> GetHash(string key, bool returnEmptyHashIfKeyNotFound = true)
        {
            if (key == null)
                return null;
            else
                key = "%" + key;

            if (allItems.ContainsKey(key))
                return (Dictionary<string, string>)allItems[key];
            if (returnEmptyHashIfKeyNotFound)
                return new Dictionary<string, string>();
            return null;
        }

        public static void SetHash(string key, Dictionary<string, string> value)
        {
            if (key == null)
                throw new ApplicationException("Key cannot be Null.");
            allItems["%" + key] = value;
        }

        public static string GetHashItem(string hashKey, string itemKey, string defaultValue = null)
        {
            if (hashKey == null || itemKey == null)
                return defaultValue;

            var hash = GetHash(hashKey, true);
            if (hash.ContainsKey(itemKey))
                return hash[itemKey];
            return defaultValue;
        }


        // support for kinda 'config flags'
        public static bool HasFlag(string key, string specificValue = null)
        {
            if (key == null)
                return false;

            key = "$" + key;
            if (!allItems.ContainsKey(key))
                return false;
            else if (specificValue == null)
                return true;
            else if (!(allItems[key] is string))
                return false;

            return (string)allItems[key] == specificValue;
        }

    }
}
