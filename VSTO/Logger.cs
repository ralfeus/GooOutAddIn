using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Net;
using System.Diagnostics;

namespace R.GoogleOutlookSync
{
    enum EventType
    {
        Debug,
        Information,
        Warning,
        Error
    }

    struct LogEntry
    {
        public DateTime date;
        public EventType type;
        public string msg;

        public LogEntry(DateTime _date, EventType _type, string _msg)
        {
            date = _date; type = _type;  msg = _msg;
        }

        public override string ToString()
        {
            return String.Format("[{0} | {1}]\t{2}\r\n", this.date, this.type, this.msg);
        }
    }

    static class Logger
    {
        private static StreamWriter _logWriter;

		public static List<LogEntry> messages = new List<LogEntry>();
		public delegate void LogUpdatedHandler(string Message);
        public static event LogUpdatedHandler LogUpdated;
        public static readonly string Folder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + System.Windows.Forms.Application.ProductName;

        static Logger()
        {
            try
            {
                if (!Directory.Exists(Folder))
                    Directory.CreateDirectory(Folder);
                _logWriter = new StreamWriter(Folder + "\\GooOut.log", true);
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle(ex);
            }
        }
    
        public static void Close()
        {
            try
            {
                if(_logWriter!=null)
                    _logWriter.Close();
            }
            catch(Exception e)
            {
                ErrorHandler.Handle(e);
            }
        }

        private static string FormatMessage(string message, EventType eventType)
        {
            return String.Format("{0}:{1}{2}", eventType, Environment.NewLine, message);
        }

		public static void Log(string message, EventType eventType)
        {
            LogEntry new_logEntry = new LogEntry(DateTime.Now, eventType, message);
#if DEBUG
            Debug.Write(new_logEntry);
#endif
            messages.Add(new_logEntry);

            try
            {
                _logWriter.Write(new_logEntry);
                _logWriter.Flush();
            }
            catch (Exception)
            {
                //ignore it, because if you handle this error, the handler will again log the message
                //ErrorHandler.Handle(ex);
            }

            //Populate LogMessage to all subscribed Logger-Outputs, but only if not Debug message, Debug messages are only logged to logfile
            if (LogUpdated != null && eventType > EventType.Debug)
                LogUpdated(new_logEntry.ToString());
        }

		public static void ClearLog()
        {
            messages.Clear();
        }

        private static void SendLog(string body, Action<bool> callback)
        {
            //var b = new StringBuilder();
            //b.AppendFormat("Application version: {0}\r\n", System.Reflection.Assembly.GetExecutingAssembly().GetName().Version);
            //b.AppendFormat("Application architecture: {0}\r\n", Utilities.GetAssemblyArchitecture());
            //b.AppendFormat("OS vesion: {0}\r\n", Environment.OSVersion);
            //b.AppendFormat("OS architecture: {0}\r\n", Utilities.GetOSArchitecture());
            //b.AppendFormat("Outlook version: {0}\r\n", Properties.Settings.Default.OutlookVersion);
            //b.AppendFormat("Outlook architecture: {0}\r\n", OutlookUtilities.GetOutlookArchitecture());
            //b.Append(body);

            //var request = (HttpWebRequest)WebRequest.Create("http://ralfeus.cti.net.ua/trac/newticket");
            //request.CookieContainer = new CookieContainer(2);
            //request.GetResponse();

            //var cookies = request.CookieContainer;
            //request = (HttpWebRequest)WebRequest.Create("http://ralfeus.cti.net.ua/trac/newticket");
            //request.Method = "POST";
            //request.ContentType = "application/x-www-form-urlencoded";
            //request.CookieContainer = cookies;

            //var stream = request.GetRequestStream();
            //var temp = "__FORM_TOKEN="+ cookies.GetCookies(new Uri("http://ralfeus.cti.net.ua/trac"))["trac_form_token"].Value + "&field_summary=Error&field_reporter=user&field_description=" + b.ToString() + "&field_type=incident&field_priority=minor&field_milestone=Reported&field_component=General&field_version=1.0&field_keywords=&field_cc=&field_owner=&submit=Create+ticket";
            //var queryBodyBytes = Encoding.UTF8.GetBytes(temp);
            //stream.Write(queryBodyBytes, 0, queryBodyBytes.Length);
            //stream.Close();

            //var result = (HttpWebResponse)request.GetResponse();
            //callback(result.StatusCode == HttpStatusCode.OK);
        }

        public static void SendSessionLog(Action<bool> callback)
        {
            var body = new StringBuilder();
            foreach (var logEntry in messages)
            {
                body.Append(logEntry.ToString());
            }
            SendLog(body.ToString(), callback);
        }

        public static void SendLogFile(Action<bool> callback)
        {
            TextReader reader = new StreamReader(Logger.Folder + "\\log.txt");
            SendLog(reader.ReadToEnd(), callback);
            reader.Close();
        }
    }
}