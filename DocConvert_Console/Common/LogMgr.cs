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
        public static void Write(string text, ConsoleColor color, LOG_LEVEL logLevel, bool noDate = false)
        {
            #region build text
            string consoleText = "";
            if (logLevel == LOG_LEVEL.ERROR)
            {
                consoleText = DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss") + " " + "[ERROR] " + text;
            }
            else if (logLevel == LOG_LEVEL.DEBUG)
            {
                consoleText = DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss") + " " + "[DEBUG] " + text;
            }
            else if (logLevel == LOG_LEVEL.INFO)
            {
                consoleText = DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss") + " " + "[INFO] " + text;
            }
            else if (logLevel == LOG_LEVEL.TRACE)
            {
                consoleText = DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss") + " " + "[TRACE] " + text;
            }
            else if (logLevel == LOG_LEVEL.WARN)
            {
                consoleText = DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss") + " " + "[WARN] " + text;
            }
            #endregion
            #region Console Log..
            Console.ForegroundColor = color;

            if(noDate)
                Console.WriteLine(text);
            else
                Console.WriteLine(consoleText);
            Console.Out.Flush();

            Console.ResetColor();
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

        public static void Write(string text, LOG_LEVEL logLevel)
        {
            if(logLevel == LOG_LEVEL.ERROR)
                Write(text, ConsoleColor.Red, logLevel);
            else
                Write(text, ConsoleColor.White, logLevel);
        }

        public static void Write(string text)
        {
            Write(text, ConsoleColor.White, LOG_LEVEL.DEBUG);
        }
    }
}
