using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Logger
{
    public sealed class LogException : ILog
    {
        private static readonly Lazy<LogException> instance = new Lazy<LogException>(() => new LogException());
        private LogException()
        {
        }
        public static LogException GetInstance
        {
            get
            {
                return instance.Value;
            }
        } 

        void ILog.LogException(string message)
        {
            String fileName = $"Exception_{DateTime.Now.ToString("yyyy-MM-dd")}.log";
            String logPath = string.Format(@"{0}\{1}", AppDomain.CurrentDomain.BaseDirectory, fileName);
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("----------------------");
            sb.AppendLine(DateTime.Now.ToString());
            sb.AppendLine(message);
            using(StreamWriter sw = new StreamWriter(logPath,true))
            {
                sw.Write(sb.ToString());
                sw.Flush();
            }
        }
    }

}
