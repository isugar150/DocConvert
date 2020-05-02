using log4net.Config;
using NLog;
using System;
using System.IO;
using System.Runtime.CompilerServices;

namespace DocConvert_Manager
{
    public enum LOG_LEVEL
    {
        TRACE,
        DEBUG,
        INFO,
        WARN,
        ERROR
    }

    public class DevLog
    {
        private static Logger logger = LogManager.GetLogger("DocConvert_Manage_Log");
        private static System.Collections.Concurrent.ConcurrentQueue<string> logMsgQueue = new System.Collections.Concurrent.ConcurrentQueue<string>();
        private static LOG_LEVEL 출력가능_로그레벨 = new LOG_LEVEL();

        static public void Init(LOG_LEVEL logLevel)
        {
            출력가능_로그레벨 = logLevel;
            String logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ManageLogConfig.xml");
            FileInfo file = new FileInfo(logPath);
            XmlConfigurator.Configure(file);
        }

        static public void ChangeLogLevel(LOG_LEVEL logLevel)
        {
            출력가능_로그레벨 = logLevel;
        }

        static public void Write(object msg, LOG_LEVEL logLevel = LOG_LEVEL.TRACE,
                                [CallerFilePath] string fileName = "",
                                [CallerMemberName] string methodName = "",
                                [CallerLineNumber] int lineNumber = 0)
        {
            if (출력가능_로그레벨 <= logLevel)
            {
                /*logMsgQueue.Enqueue(string.Format("{0}   Method: {1}   Message: {2}", System.DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss.fff"), methodName, msg));*/
                logMsgQueue.Enqueue(string.Format("{0}   Message: {1}", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff"), msg));
            }
            if (logLevel == LOG_LEVEL.ERROR)
            {
                logger.Error(msg.ToString());
            }
            else if (logLevel == LOG_LEVEL.DEBUG)
            {
                logger.Debug(msg.ToString());
            }
            else if (logLevel == LOG_LEVEL.INFO)
            {
                logger.Info(msg.ToString());
            }
            else if (logLevel == LOG_LEVEL.TRACE)
            {
                logger.Trace(msg.ToString());
            }
            else if (logLevel == LOG_LEVEL.WARN)
            {
                logger.Warn(msg.ToString());
            }
        }

        static public bool GetLog(out string msg)
        {
            if (logMsgQueue.TryDequeue(out msg))
            {
                return true;
            }

            return false;
        }

    }
}
