using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Common.Logging;
using Common.Logging.Log4Net;

namespace RMA_Roots
{
    public class Log
    {
        static ILog log = null;

        static Log()
        {
            log = LogManager.GetLogger(typeof(Log));
        }

        public static void Info(object msg)
        {
            log.Info(msg);
        }
        public static void Warn(object msg)
        {
            log.Warn(msg);
        }
        public static void Fatal(object msg, Exception exception)
        {
            log.Fatal(msg, exception);
        }
    }
}
