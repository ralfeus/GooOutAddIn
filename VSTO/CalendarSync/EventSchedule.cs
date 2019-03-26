using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Calendar.v3;

namespace R.GoogleOutlookSync
{
    class EventSchedule
    {
        internal bool AllDayEvent { get; set; }
        internal DateTime EndTime { get; set; }
        internal EventRecurrence RecurrencePattern { get; set; }
        internal DateTime StartTime { get; set; }
        internal string TimeZone { get; set; }

        internal EventSchedule(Event googleItem)
        {
            this.AllDayEvent = !string.IsNullOrEmpty(googleItem.Start.Date);
            this.TimeZone = googleItem.Start.TimeZone;
            if (this.AllDayEvent)
            {
                this.StartTime = DateTime.Parse(googleItem.Start.Date);
                this.EndTime = DateTime.Parse(googleItem.End.Date);
            }
            else
            {
                this.StartTime = googleItem.Start.DateTime.Value;
                this.EndTime = googleItem.End.DateTime.Value;
            }
        }

        internal EventSchedule(Event googleItem, Calendar calendar, CalendarService calendarService)
            :this(googleItem)
        { 
            if (googleItem.Recurrence != null)
            {
                this.RecurrencePattern = new EventRecurrence(googleItem, calendar, calendarService);
            }
        }

        //internal EventSchedule(Event googleItem, List<Event> recurrenceExceptions = null)
        //{
        //    this.AllDayEvent = !string.IsNullOrEmpty(googleItem.Start.Date);
        //    this.TimeZone = googleItem.Start.TimeZone;
        //    if (this.AllDayEvent)
        //    {
        //        this.StartTime = DateTime.Parse(googleItem.Start.Date);
        //        this.EndTime = DateTime.Parse(googleItem.End.Date);
        //    } else
        //    {
        //        this.StartTime = googleItem.Start.DateTime.Value;
        //        this.EndTime = googleItem.End.DateTime.Value;
        //    }
        //    if (googleItem.Recurrence != null)
        //    {
        //        this.RecurrencePattern = new EventRecurrence(googleItem, recurrenceExceptions);
        //    }
        //}

        internal EventSchedule(AppointmentItem outlookItem)
        {
            this.StartTime = outlookItem.Start;
            this.EndTime = outlookItem.End;
            this.AllDayEvent = outlookItem.AllDayEvent;
            this.TimeZone = outlookItem.StartTimeZone.ID;
            if (outlookItem.IsRecurring && (outlookItem.RecurrenceState == OlRecurrenceState.olApptMaster))
                this.RecurrencePattern = new EventRecurrence(outlookItem.GetRecurrencePattern());
        }

        public static bool operator ==(EventSchedule source, EventSchedule target)
        {
            return source.Equals(target);
        }

        public static bool operator !=(EventSchedule source, EventSchedule target)
        {
            return !source.Equals(target);
        }

        public override bool Equals(object obj)
        {
            var target = (EventSchedule)obj;
            return
                (this.StartTime == target.StartTime) &&
                (this.EndTime == target.EndTime) &&
                (this.AllDayEvent == target.AllDayEvent) &&
                (this.RecurrencePattern == target.RecurrencePattern);
        }

        public override int GetHashCode()
        {
            return this.StartTime.GetHashCode() ^ this.EndTime.GetHashCode();
        }

        /// <summary>
        /// Generates Google recurrence pattern
        /// Because Google recurrence pattern contains scheduling information (like start time and end time) 
        /// it's better to delegate such work to EventSchedule object
        /// </summary>
        internal IList<string> GetGoogleRecurrence()
        {
            string result = "";
            /// This part isn't needed for new Google API
            //if (this.AllDayEvent)
            //{
            //    result = String.Format("DTSTART;VALUE=DATE:{0}\r\n", this.StartTime.ToString("yyyyMMdd"));
            //    result += String.Format("DTEND;VALUE=DATE:{0}\r\n", this.EndTime.ToString("yyyyMMdd"));
            //}
            //else
            //{
            //    result = String.Format("DTSTART;TZID={0}:{1}\r\n", TimeZones.GetTZ(this.TimeZone), this.StartTime.ToString("yyyyMMddTHHmmss"));
            //    result += String.Format("DTEND;TZID={0}:{1}\r\n", TimeZones.GetTZ(this.TimeZone), this.EndTime.ToString("yyyyMMddTHHmmss"));
            //}

            /// RFC 2445 tells VTIMEZONE must be set if event start is defined by DTSTART and DTEND elements
            //var timeZone = String.Format("VTIMEZONE={0};", "");

            /// It's necessary to define week interval in such way in order to add week interval
            /// (if any) to BYDAY element
            var instance = this.RecurrencePattern.WeekInterval != 0 ? this.RecurrencePattern.WeekInterval.ToString() : "";
            var byDay = this.RecurrencePattern.DayOfWeekMask != 0 ? String.Format("BYDAY={0};", Regex.Replace(this.RecurrencePattern.DayOfWeekMask.ToString(), "((\\w{2})\\w+)", instance + "$2").ToUpper().Replace(" ", "")) : "";
            var byMonthDay = this.RecurrencePattern.DayOfMonth != 0 ? String.Format("BYMONTHDAY={0};", this.RecurrencePattern.DayOfMonth) : "";
            var count = this.RecurrencePattern.Count != 0 ? String.Format("COUNT={0};", this.RecurrencePattern.Count) : "";
            var frequency = String.Format("FREQ={0};", this.RecurrencePattern.Frequency.ToString().ToUpper());
            var interval = this.RecurrencePattern.Interval > 1 ? String.Format("INTERVAL={0};", this.RecurrencePattern.Interval) : "";
            var month = this.RecurrencePattern.Frequency == RecurrenceFrequency.Yearly ? String.Format("BYMONTH={0};", this.RecurrencePattern.Month) : "";
            var until = this.RecurrencePattern.EndMethod == EndBy.Date ? String.Format("UNTIL={0}", this.RecurrencePattern.EndDate.ToString("yyyyMMddTHHmmssZ")) : "";

            var rule = string.Format("RRULE:{0}{1}{2}{3}{4}{5}{6}", frequency, interval, count, month, byDay, byMonthDay, until);
            result += rule.Substring(0, rule.Length - 1);

            //TODO: Add recurrence exceptions

            return new List<string> { result };
        }

        internal void GetOutlookRecurrence(AppointmentItem outlookItem)
        {
            RecurrencePattern outlookRec = outlookItem.GetRecurrencePattern();
            if (!this.AllDayEvent)
            {
                /// Set time of the event taking consideration difference of time zones. 
                /// This makes sense only in case the event isn't all day one
                /// Google event time is already converted to local time zone
                var googleTimeZone = TimeZoneInfo.Local; 
                var outlookTimeZone = TimeZoneInfo.FindSystemTimeZoneById(outlookItem.StartTimeZone.ID);
                /// StartTime property contains only time (not date)
                outlookRec.StartTime = TimeZoneInfo.ConvertTime(this.StartTime, googleTimeZone, outlookTimeZone);
            }
 
            outlookRec.Duration = (int)(this.EndTime - this.StartTime).TotalMinutes;
            outlookRec.RecurrenceType = (OlRecurrenceType)this.RecurrencePattern.Frequency;
            if ((this.RecurrencePattern.Frequency == RecurrenceFrequency.Weekly) || (this.RecurrencePattern.Frequency == RecurrenceFrequency.MonthlyNth))
                outlookRec.DayOfWeekMask = (OlDaysOfWeek)this.RecurrencePattern.DayOfWeekMask;
            if (this.RecurrencePattern.Interval > 1)
                if ((outlookRec.RecurrenceType != OlRecurrenceType.olRecursYearly) || (VSTO.Properties.Settings.Default.OutlookVersion >= 14))
                    outlookRec.Interval = this.RecurrencePattern.Interval;
            outlookRec.PatternStartDate = this.RecurrencePattern.StartDate;
            if (this.RecurrencePattern.EndMethod == EndBy.Date)
                outlookRec.PatternEndDate = this.RecurrencePattern.EndDate;
            if (this.RecurrencePattern.EndMethod == EndBy.OccurencesCount)
                outlookRec.Occurrences = this.RecurrencePattern.Count;
            if ((this.RecurrencePattern.Frequency == RecurrenceFrequency.Monthly) || (this.RecurrencePattern.Frequency == RecurrenceFrequency.Yearly))
                outlookRec.DayOfMonth = this.RecurrencePattern.DayOfMonth;
            /// As it was found out RecurrencePattern.Instance is valid for olRecursMonthNth only. 
            /// Not sure yet whether Google provides such frequency model. If not - it will be just omited
            //outlookRec.Instance = this.WeekInterval;
            if (this.RecurrencePattern.Frequency == RecurrenceFrequency.Yearly)
                outlookRec.MonthOfYear = this.RecurrencePattern.Month;
            //outlookRec.NoEndDate = this.EndMethod == EndBy.NoEnd;

        }

        internal void ToGoogle(Event googleItem)
        {
            if (this.RecurrencePattern != null)
            {
                googleItem.Recurrence = this.GetGoogleRecurrence();
            }
            else
            {
                if (this.AllDayEvent)
                {
                    googleItem.Start.Date = this.StartTime.ToString("yyyy-MM-dd");
                    googleItem.End.Date = this.EndTime.ToString("yyyy-MM-dd");
                } else
                {
                    googleItem.Start.DateTime = this.StartTime;
                    googleItem.End.DateTime = this.EndTime;
                }
            }
        }

        internal void ToOutlook(AppointmentItem outlookItem)
        {
            
        }
    }
}
