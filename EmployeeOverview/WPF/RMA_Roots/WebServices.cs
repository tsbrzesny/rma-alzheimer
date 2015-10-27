using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Xml;

namespace RMA_Roots
{
    /// <summary>
    /// Client Area: https://www.iban.net/clients/
    /// Username: runmyaccounts
    /// Password: 0b84ac34
    /// (access mail: 21.5.2015, Thomas Brändle)
    /// </summary>
    public class IBAN_Net
    {
        static string ibanRoot = AppConfig.GetItem("IbanEndpoint");
        static string ibanAPIKey = AppConfig.GetItem("IbanAPIKey");

        public class IBAN_Net_Result
        {
            public string iban { get; set; }
            public string bank_name { get; set; }
            public string country { get; set; }
            public string city { get; set; }
            public string address { get; set; }
            public string swift { get; set; }
            public bool? sepa { get; set; }
            public bool valid { get; set; }
            public string error { get; set; }
        }


        public static IBAN_Net_Result ReverseIBAN(string iban)
        {
            if (iban == null)
                return null;

            var endpoint = string.Format(ibanRoot + "?api_key={0}&iban={1}", ibanAPIKey, iban);
            var rStr = WebServices.getServiceResult(endpoint);

            try 
	        {	        
		        var xDoc = new XmlDocument();
                xDoc.LoadXml(rStr);
                var json = JsonConvert.SerializeXmlNode(xDoc, Newtonsoft.Json.Formatting.None, true);
                return JsonConvert.DeserializeObject<IBAN_Net_Result>(json.ToString());
            }

	        catch (Exception)
            {
		        throw new ApplicationException("IBAN_Net.ReverseIBAN: invalid service response.");
            }
        }

    }


    public class Slack_API
    {

        // beno's token: xoxp-3569429651-5010389533-5146963048-ecd20d
        public const string AccessToken = "xoxp-3569429651-5010389533-5146963048-ecd20d";


        public class Slack_Channel
        {
            public string id { get; set; }
            public string name { get; set; }
            public bool is_archived { get; set; }

            /*
             * example & other members:
                {
	                "id":"C0521HWP0",
	                "name":"banken_schnittstellen",
	                "is_channel":true,
	                "created":1432711733,
	                "creator":"U04P5GAF0",
	                "is_archived":false,
	                "is_general":false,
	                "is_member":false,
	                "members":["U03GRCMKF","U04P5AP72","U04P5GAF0","U04V3M8N9","U050ABFFP"],
	                "topic":{"value":"","creator":"","last_set":0},
	                "purpose":{"value":"Erkundung und Definition m\u00f6glicher Banken-Schnittstelle","creator":"U04P5GAF0","last_set":1432711733},
	                "num_members":5
                }
             */

            public static Slack_Channel FindByName(string name)
            {
                if (name == null)
                    return null;

                // load channel list & find exact name
                var lsc = Slack_API.ChannelList();
                var result = lsc.FirstOrDefault(sc => sc.name == name);
                if (result != null)
                    return result;

                // check for a unique 'contains' match..
                var matches = lsc.FindAll(sc => sc.name.Contains(name.ToLower()));
                if (matches.Count() == 1)
                    return matches[0];

                return null;
            }

            public PostMessage_Rsp PostMessage(string message, string username = null, bool link_names = false)
            {
                return Slack_API.PostMessage(id, message, username, link_names);
            }
        }

        public static List<Slack_Channel> ChannelList(bool excludeArchivedChannels = true)
        {
            // example:
            // https://slack.com/api/channels.list?token=xoxp-3569429651-4569769358-5146818942-8c58c7&exclude_archived=1

            var endpoint = string.Format("https://slack.com/api/channels.list?token={0}", AccessToken);
            if (excludeArchivedChannels)
                endpoint += "&exclude_archived=1";

            string responseStr = null;
            try
            {
                responseStr = WebServices.getServiceResult(endpoint);

                dynamic jo = JToken.Parse(responseStr);
                var ok = (bool)jo.ok;
                var channels = jo.channels as JArray;

                if (!ok)
                    throw new ApplicationException(jo.error);
                if (channels == null)
                    throw new ApplicationException("Channel data not accessible!");

                var lsc = new List<Slack_Channel>();
                JsonConvert.PopulateObject(channels.ToString(), lsc);
                return lsc;
            }

            catch (Exception e)
            {
                throw new ApplicationException("Slack_API.ChannelList: " + e.Message, e);
            }
        }



        public class PostMessage_Rsp
        {
            // see: https://api.slack.com/methods/chat.postMessage
            public bool ok { get; set; }
            public string ts { get; set; }
            public string channel { get; set; }
            public JObject message { get; set; }
            public string error { get; set; }

            public string UpdateMessage(string text)
            {
                var endpoint = string.Format("https://slack.com/api/chat.update?token={0}&ts={1}&channel={2}&text={3}",
                                             AccessToken, ts, channel, WebServices.UrlEncode(text));
                var responseStr = WebServices.getServiceResult(endpoint);

                dynamic jo = JToken.Parse(responseStr);
                if ((bool)jo.ok)
                    return null;
                else
                    return jo.error;
            }

            public string DeleteMessage()
            {
                var endpoint = string.Format("https://slack.com/api/chat.delete?token={0}&ts={1}&channel={2}", AccessToken, ts, channel);
                var responseStr = WebServices.getServiceResult(endpoint);

                dynamic jo = JToken.Parse(responseStr);
                if ((bool)jo.ok)
                    return null;
                else
                    return jo.error;
            }

        }
        //
        public static PostMessage_Rsp PostMessage(string channelId, string message, string username = null, bool link_names = false)
        {
            // see: https://api.slack.com/methods/chat.postMessage
            // example:
            // https://slack.com/api/chat.postMessage?token=xoxp-3569429651-5010389533-5146963048-ecd20d&channel=C053600BQ&text=hello+world&username=FTX

            var endpoint = string.Format("https://slack.com/api/chat.postMessage?token={0}&channel={1}&text={2}",
                                         AccessToken, channelId, WebServices.UrlEncode(message));
            if (username != null)
                endpoint += "&username=" + WebServices.UrlEncode(username);
            if (link_names)
                endpoint += "&link_names=1";

            string responseStr = null;
            try
            {
                responseStr = WebServices.getServiceResult(endpoint);
                var pmr = new PostMessage_Rsp();
                JsonConvert.PopulateObject(responseStr, pmr);
                return pmr;
            }

            catch (Exception)
            {
                throw;
            }
        }

    }

    
    public class WebServices
    {

        public static string getServiceResult(string serviceUrl) {
    	    HttpWebRequest HttpWReq;
    	    HttpWebResponse HttpWResp;
    	    HttpWReq = (HttpWebRequest)WebRequest.Create(serviceUrl);
    	    HttpWReq.Method = "GET";
    	    HttpWResp = (HttpWebResponse)HttpWReq.GetResponse();
    	    if (HttpWResp.StatusCode == HttpStatusCode.OK)
    	    {
    		    using (var reader = new StreamReader(HttpWResp.GetResponseStream()))
                    return reader.ReadToEnd();
    	    }

    	    else
    		    throw new ApplicationException("WebServices.getServiceResult error: "+ HttpWResp.StatusCode.ToString());
        }

        public static string UrlEncode(string s)
        {
            // replaces all URL special characters by their URL equivalent

            if (s != null)
                s = Regex.Replace(s, @"[^a-zA-Z0-9\-_.~]", UrlEncode_UrlReplacement);
            return s;
        }

        private static string UrlEncode_UrlReplacement(Match m)
        {
            var result = "";
            foreach (var b in System.Text.Encoding.UTF8.GetBytes(m.Groups[0].Value))
                result += string.Format("%{0:X2}", b);
            return result;
        }

    }


}
