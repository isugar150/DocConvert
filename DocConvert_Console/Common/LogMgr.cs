using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocConvert_Console.Common
{
    public enum LOG_LEVEL
    {
        TRACE,
        DEBUG,
        INFO,
        WARN,
        ERROR
    }
    public class LogMgr
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Console_Log");
        public static void Write(string text, ConsoleColor color, LOG_LEVEL logLevel)
        {
            #region Console Log..
            ConsoleColor originalColor = Console.ForegroundColor;
            Console.ForegroundColor = color;

            Console.WriteLine(text);
            Console.Out.Flush();

            Console.ForegroundColor = originalColor;
            #endregion
            #region NLog..

            if (logLevel == LOG_LEVEL.ERROR)
            {
                logger.Error(text);
            }
            else if (logLevel == LOG_LEVEL.DEBUG)
            {
                logger.Debug(text);
            }
            else if (logLevel == LOG_LEVEL.INFO)
            {
                logger.Info(text);
            }
            else if (logLevel == LOG_LEVEL.TRACE)
            {
                logger.Trace(text);
            }
            else if (logLevel == LOG_LEVEL.WARN)
            {
                logger.Warn(text);
            }
            #endregion
        }
    }
}
