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
        ERROR,
        FATAL
    }
    public class LogMgr
    {
        private static Logger consoleLogger = LogManager.GetLogger("DocConvert_Console_Log");
        private static Logger coreLogger = LogManager.GetLogger("DocConvert_Core_Log");

        /// <summary>
        /// 로그를 작성합니다.
        /// 기본 컬러: White
        /// 기본 레벨: Debug
        /// </summary>
        /// <param name="text">작성할 텍스트</param>
        /// <param name="color">콘솔 출력 컬러</param>
        /// <param name="logLevel">로그 레벨</param>
        /// <param name="noDate">콘솔에 날짜와 로그레벨을 출력하지 않습니다.</param>
        public static void Write(string text, ConsoleColor color, LOG_LEVEL logLevel, bool noDate = false)
        {
            #region build text
            string consoleText = "";
            if (logLevel == LOG_LEVEL.FATAL)
            {
                consoleText = DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss") + " " + "[FATAL] " + text;
            }
            else if(logLevel == LOG_LEVEL.ERROR)
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
            if (logLevel == LOG_LEVEL.FATAL)
            {
                consoleLogger.Fatal(text);
            }
            else if (logLevel == LOG_LEVEL.ERROR)
            {
                consoleLogger.Error(text);
            }
            else if (logLevel == LOG_LEVEL.DEBUG)
            {
                consoleLogger.Debug(text);
            }
            else if (logLevel == LOG_LEVEL.INFO)
            {
                consoleLogger.Info(text);
            }
            else if (logLevel == LOG_LEVEL.TRACE)
            {
                consoleLogger.Trace(text);
            }
            else if (logLevel == LOG_LEVEL.WARN)
            {
                consoleLogger.Warn(text);
            }
            #endregion
        }

        /// <summary>
        /// 로그를 작성합니다.
        /// 기본 컬러: White
        /// 기본 레벨: Debug
        /// </summary>
        /// <param name="text">작성할 텍스트</param>
        /// <param name="logLevel">로그 레벨</param>
        public static void Write(string text, LOG_LEVEL logLevel)
        {
            if(logLevel == LOG_LEVEL.ERROR)
                Write(text, ConsoleColor.Red, logLevel);
            else
                Write(text, ConsoleColor.White, logLevel);
        }

        /// <summary>
        /// 로그를 작성합니다.
        /// 기본 컬러: White
        /// 기본 레벨: Debug
        /// </summary>
        /// <param name="text">작성할 텍스트</param>
        public static void Write(string text)
        {
            Write(text, ConsoleColor.White, LOG_LEVEL.DEBUG);
        }

        public static void Write(string text, ConsoleColor consoleColor)
        {
            Write(text, consoleColor, LOG_LEVEL.DEBUG);
        }

        public static string getLogLevel(string loggerName)
        {
            string logLevel = "";
            if (loggerName.Equals("DocConvert_Core_Log"))
            {
                if (coreLogger.IsTraceEnabled)
                    logLevel = "TRACE";
                else if (coreLogger.IsDebugEnabled)
                    logLevel = "DEBUG";
                else if (coreLogger.IsInfoEnabled)
                    logLevel = "INFO";
                else if (coreLogger.IsWarnEnabled)
                    logLevel = "WARN";
                else if (coreLogger.IsErrorEnabled)
                    logLevel = "ERROR";
                else logLevel = null;
            } else if (loggerName.Equals("DocConvert_Console_Log"))
            {
                if (consoleLogger.IsTraceEnabled)
                    logLevel = "TRACE";
                else if (consoleLogger.IsDebugEnabled)
                    logLevel = "DEBUG";
                else if (consoleLogger.IsInfoEnabled)
                    logLevel = "INFO";
                else if (consoleLogger.IsWarnEnabled)
                    logLevel = "WARN";
                else if (consoleLogger.IsErrorEnabled)
                    logLevel = "ERROR";
                else logLevel = null;
            }
            else
            {
                logLevel = null;
            }

            return logLevel;
        }
    }
}
