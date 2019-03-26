using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.IO;
using System.Xml;

namespace R.GoogleOutlookSync
{
    class TimeZones
    {
        private Dictionary<string, string> _timeZonesMap = null;
        private static TimeZones _instance;

        private static TimeZones Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new TimeZones();
                return _instance;
            }
        }

        private TimeZones()
        {
            this._timeZonesMap = new Dictionary<string,string>();
            var xmlStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("VSTO.CalendarSync.TimeZones.xml");
            var reader = XmlReader.Create(xmlStream);
            var tz = "";
            var windowsTimeZone = "";
            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                        switch (reader.Name)
                        {
                            case "TZ":
                                reader.Read();
                                tz = reader.Value;
                                break;
                            case "WindowsTimeZone":
                                reader.Read();
                                windowsTimeZone = reader.Value;
                                break;
                        }
                        break;
                    case XmlNodeType.EndElement:
                        switch (reader.Name)
                        {
                            case "Zone":
                                this._timeZonesMap.Add(tz, windowsTimeZone);
                                break;
                        }
                        break;
                }
            }
            reader.Close();
            xmlStream.Close();
        }

        public static string GetTZ(string windowsTimeZone)
        {
            try
            {
                return Instance._timeZonesMap.First(zone => zone.Value == windowsTimeZone).Key;
            }
            catch (Exception exc)
            {
                Logger.Log(String.Format("Couldn't find Windows time zone '{0}' in Google", windowsTimeZone), EventType.Error);
                throw exc;
            }
        }

        public static string GetWindowsTimeZone(string tz)
        {
            try
            {
                return Instance._timeZonesMap[tz];
            }
            catch (Exception exc)
            {
                Logger.Log(String.Format("Couldn't find Google time zone '{0}' in Windows", tz), EventType.Error);
                throw exc;
            }
        }
    }
}
