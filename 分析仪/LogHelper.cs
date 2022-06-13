using System;

public class LogHelp
{
    public class LogHelper
    {
        //这里的 loginfo 和 log4net.config 里的名字要一样
        public static readonly log4net.ILog log_info = log4net.LogManager.GetLogger("loginfo");

        //这里的 logerror 和 log4net.config 里的名字要一样
        public static readonly log4net.ILog log_error = log4net.LogManager.GetLogger("logerror");

        public static void WriteLog(string info)
        {
            if (log_info.IsInfoEnabled)
            {
                log_info.Info(info);
            }
        }

        public static void WriteLog(string error, Exception ex)
        {
            if (log_error.IsErrorEnabled)
            {
                log_error.Error(error, ex);
            }
        }
    }
}
