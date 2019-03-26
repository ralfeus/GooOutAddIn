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
    class EventRecurrence
    {

        //private static Regex _startTimePattern = new Regex("DTSTART(;TZID=.+?)?(;VALUE=DATE)?:(\\d{4})(\\d{2})(\\d{2})(T(\\d{2})(\\d{2})(\\d{2})Z?)?(\r|$)", RegexOptions.Singleline | RegexOptions.Compiled);
        //private static Regex _endTimePattern = new Regex("DTEND(;TZID=.+?)?(;VALUE=DATE)?:(\\d{4})(\\d{2})(\\d{2})(T(\\d{2})(\\d{2})(\\d{2})Z?)?(\r|$)", RegexOptions.Singleline | RegexOptions.Compiled);
        private static Regex _frequencyPattern = new Regex("FREQ=(.*?)(;|\r|$)", RegexOptions.Singleline | RegexOptions.Compiled);
        private static Regex _byDayPattern = new Regex("(?<!BEGIN.*)BYDAY=(((-?\\d*)([A-Z]{2},?))+)(;|\r|$)", RegexOptions.Singleline | RegexOptions.Compiled);
        private static Regex _byMonthPattern = new Regex("(?<!BEGIN.*)BYMONTH=(\\d+)(;|\r|$)", RegexOptions.Singleline | RegexOptions.Compiled);
        private static Regex _byMonthDayPattern = new Regex("(?<!BEGIN.*)BYMONTHDAY=(\\d+)(;|\r|$)", RegexOptions.Singleline | RegexOptions.Compiled);
        private static Regex _countPattern = new Regex("COUNT=(\\d+)(;|\r|$)", RegexOptions.Singleline | RegexOptions.Compiled);
        private static Regex _intervalPattern = new Regex("INTERVAL=(\\d+)(;|\r|$)", RegexOptions.Singleline | RegexOptions.Compiled);
        /// <summary>
        /// Defines regexp for matching recurrence end date.
        /// According to RFC 2445 it can be also time. But neither Google nor Outlook define time for end of recurrence.
        /// So we assume UNTIL element will contain date only
        /// </summary>
        private static Regex _untilPattern = new Regex("UNTIL=(\\d{4})(\\d{2})(\\d{2})(;|\r|$)", RegexOptions.Compiled | RegexOptions.Singleline); // (T(\\d{2})(\\d{2})(\\d{2}))? - for time (if necessary)

        /// <summary> 6j 
        /// Count of recurrence instances
        /// </summary>
        internal int Count { get; set; }
        /// <summary>
        /// Defines in which day of month even reoccurs
        /// </summary>
        internal int DayOfMonth { get; set; }
        /// <summary>
        /// Defines in which days of week event reoccurs
        /// </summary>
        internal DayOfWeek DayOfWeekMask { get; set; }
        /// <summary>
        /// Last day when event should reoccur
        /// </summary>
        internal DateTime EndDate { get; set; }
        /// <summary>
        /// Specifies how recurrence ends
        /// </summary>
        internal EndBy EndMethod { get; set; }
        /// <summary>
        /// Frequency of recurrence
        /// </summary>
        internal RecurrenceFrequency Frequency { get; set; }
        /// <summary>
        /// Interval between two recurrence frequency elements
        /// Example: Interval = 2 and Frequency = Weekly means every second week
        /// </summary>
        internal int Interval { get; set; }
        /// <summary>
        /// Defines in which month even reoccurs
        /// </summary>
        internal int Month { get; set; }
        /// <summary>
        /// Defines date when recurrence starts. Usually it's date of event start
        /// </summary>
        internal DateTime StartDate { get; set; }
        /// <summary>
        /// Defines week interval for monthly and yearly recurrences.
        /// Example: Every 2nd Tuesday of month (monthly) - WeekInterval = 2
        /// Example: Every 3rd Monday of every February (yearly) - WeekInterval = 3
        /// </summary>
        internal int WeekInterval { get; set; }
        internal IList<RecurrenceException> Exceptions { get; set; }

        public EventRecurrence()
        {
            this.Exceptions = new List<RecurrenceException>();
        }

        public EventRecurrence(Event googleItem, Calendar calendar, CalendarService calendarService)
            :this(googleItem)
        {
            /// Initialize recurrence exceptions
            var instancesRequest = calendarService.Events.Instances(calendar.Id, googleItem.Id);
            instancesRequest.ShowDeleted = true;
            var result = instancesRequest.Execute();
            var instances = result.Items;
            var defaultReminders = result.DefaultReminders; 
            var exceptions = instances.Where(e => e.Status == "cancelled" || e.Start.DateTime != e.OriginalStartTime.DateTime);
            if (exceptions.Count() > 0)
            {
                this.Exceptions = exceptions.Select(exception => new RecurrenceException(exception, defaultReminders)).ToList();
            }
        }

        public EventRecurrence(Event googleItem)
            :this()
        {
            if (googleItem.Recurrence == null)
                throw new ArgumentNullException("googleRecurrence", "Recurrence pattern is null");
            var recPattern = string.Join(@"\n", googleItem.Recurrence);
            this.StartDate = (googleItem.Start.DateTime ?? DateTime.Parse(googleItem.Start.Date)).Date;
            var endDateMatch = _untilPattern.Match(recPattern);
            if (endDateMatch.Success)
                this.EndDate = new DateTime(
                    Convert.ToInt32(endDateMatch.Groups[1].Value),
                    Convert.ToInt32(endDateMatch.Groups[2].Value),
                    Convert.ToInt32(endDateMatch.Groups[3].Value));
            else
                this.EndDate = DateTime.MaxValue;
            this.Frequency = (RecurrenceFrequency)Enum.Parse(typeof(RecurrenceFrequency), _frequencyPattern.Match(recPattern).Groups[1].Value, true);
            var intervalMatch = _intervalPattern.Match(recPattern);
            this.Interval = intervalMatch.Success ? Convert.ToInt32(intervalMatch.Groups[1].Value) : 1;
            var countMatch = _countPattern.Match(recPattern);
            this.Count = countMatch.Success ? Convert.ToInt32(countMatch.Groups[1].Value) : 0;
            if (this.Count != 0)
                this.EndMethod = EndBy.OccurencesCount;
            else if (this.EndDate == DateTime.MaxValue)
                this.EndMethod = EndBy.NoEnd;
            else
                this.EndMethod = EndBy.Date;
            if (this.Frequency == RecurrenceFrequency.Weekly)
                this.DayOfWeekMask = GetDayOfWeekMask(_byDayPattern.Match(recPattern).Groups[1].Value);
            else if (this.Frequency == RecurrenceFrequency.Monthly)
            {
                var byMonthDayMatch = _byMonthDayPattern.Match(recPattern);
                this.DayOfMonth = byMonthDayMatch.Success ? Convert.ToInt32(byMonthDayMatch.Groups[1].Value) : this.StartDate.Day;
                var weekIntervalMatch = _byDayPattern.Match(recPattern).Groups[3];
                this.WeekInterval = weekIntervalMatch.Success ? Convert.ToInt32(weekIntervalMatch.Value) : 0;
                if (this.WeekInterval != 0)
                    this.Frequency = RecurrenceFrequency.MonthlyNth;
                this.DayOfWeekMask = GetDayOfWeekMask(_byDayPattern.Match(recPattern).Groups[1].Value);
            }
            else if (this.Frequency == RecurrenceFrequency.Yearly)
            {
                var byMonthDayMatch = _byMonthDayPattern.Match(recPattern);
                /// Some Google entries (at least yearly ones) don't have BYMONTH and BYMONTHDAY elements
                /// In this case it seems day and month are taken from start date
                if (byMonthDayMatch.Success)
                    this.DayOfMonth = Convert.ToInt32(byMonthDayMatch.Groups[1].Value);
                else
                    this.DayOfMonth = this.StartDate.Day;
                var byMonthMatch = _byMonthPattern.Match(recPattern);
                if (byMonthMatch.Success)
                    this.Month = Convert.ToInt32(byMonthMatch.Groups[1].Value);
                else
                    this.Month = this.StartDate.Month;
                var weekIntervalMatch = _byDayPattern.Match(recPattern).Groups[3];
                this.WeekInterval = weekIntervalMatch.Success ? Convert.ToInt32(weekIntervalMatch.Value) : 0;
                this.DayOfWeekMask = GetDayOfWeekMask(_byDayPattern.Match(recPattern).Groups[1].Value);
                /// Outlook's calendar model doesn't support intervals for yearly recurrence patterns for Outlook version before 2010.
                if ((this.Interval != 1) && (VSTO.Properties.Settings.Default.OutlookVersion < 14))
                    throw new IncompatibleRecurrencePatternException(Target.Google, VSTO.Properties.Resources.Error_WrongYearlyInterval);
            }
        }

        public EventRecurrence(RecurrencePattern outlookRecurrence)
            :this()
        {
            if (outlookRecurrence == null)
                throw new ArgumentNullException("outlookRecurrence", "Recurrence pattern is null");

            switch (outlookRecurrence.RecurrenceType)
            {
                case OlRecurrenceType.olRecursDaily:
                    this.Frequency = RecurrenceFrequency.Daily;
                    break;
                case OlRecurrenceType.olRecursWeekly:
                    this.Frequency = RecurrenceFrequency.Weekly;
                    break;
                case OlRecurrenceType.olRecursMonthly:
                case OlRecurrenceType.olRecursMonthNth:
                    this.Frequency = RecurrenceFrequency.Monthly;
                    break;
                case OlRecurrenceType.olRecursYearly:
                case OlRecurrenceType.olRecursYearNth:
                    this.Frequency = RecurrenceFrequency.Yearly;
                    break;
            }
            this.Interval = outlookRecurrence.Interval;
            if (outlookRecurrence.NoEndDate)
                this.EndMethod = EndBy.NoEnd;
            else if (outlookRecurrence.Occurrences != 0)
            {
                this.EndMethod = EndBy.OccurencesCount;
                this.Count = outlookRecurrence.Occurrences;
            }
            else
                this.EndMethod = EndBy.Date;
            this.EndDate = this.EndMethod != EndBy.Date ? DateTime.MaxValue : outlookRecurrence.PatternEndDate;
            this.StartDate = outlookRecurrence.PatternStartDate;
            if (this.Frequency == RecurrenceFrequency.Weekly)
                this.DayOfWeekMask = (DayOfWeek)outlookRecurrence.DayOfWeekMask;
            else if (this.Frequency == RecurrenceFrequency.Monthly)
            {
                this.DayOfMonth = outlookRecurrence.DayOfMonth;
                this.WeekInterval = outlookRecurrence.Instance;
                this.DayOfWeekMask = (DayOfWeek)outlookRecurrence.DayOfWeekMask;
            }
            else if (this.Frequency == RecurrenceFrequency.Yearly)
            {
                this.DayOfMonth = outlookRecurrence.DayOfMonth;
                this.Month = outlookRecurrence.MonthOfYear;
                this.WeekInterval = outlookRecurrence.Instance;
                this.DayOfWeekMask = (DayOfWeek)outlookRecurrence.DayOfWeekMask;
                this.Interval = 1;
            }

            /// Not sure yet whether it's necessary
            var exceptions = outlookRecurrence.Exceptions;
            try
            {
                if (exceptions.Count > 0)
                {
                    //this.Exceptions = new List<RecurrenceException>(exceptions.Count);
                    foreach (Microsoft.Office.Interop.Outlook.Exception outlookException in exceptions)
                    {
                        try {
                            var exception = new RecurrenceException(this, outlookException.OriginalDate) {
                                Deleted = outlookException.Deleted,
                                ModifiedEvent = outlookException.Deleted ? null : new CalendarEvent(outlookException.AppointmentItem)
                            };
                            this.Exceptions.Add(exception);
                        } catch (COMException exc) {
                            if (exc.HResult == -1802485755) {
                                Logger.Log(string.Format(
                                "The exception of '{0}' with original start time {1} couldn't be read. Ignoring this exception",
                                outlookException.Parent.Subject, outlookException.OriginalDate), EventType.Warning);
                            }
                        }
                    }
                }
            }
            finally
            {
                if (exceptions != null)
                    Marshal.ReleaseComObject(exceptions);
            }
        }

        public static bool operator ==(EventRecurrence source, EventRecurrence target)
        {
            /// Cast source and target to Object in order to avoid loop
            if ((object)source == null)
                return (object)target == null;
            else
                return 
                    ((object)target != null) && source.Equals(target);
        }

        public static bool operator !=(EventRecurrence source, EventRecurrence target)
        {
            /// Cast source and target to Object in order to avoid loop
            if ((object)source == null)
                return (object)target != null;
            else
                return (object)target == null || !source.Equals(target);
        }

        public override bool Equals(object obj)
        {
            if (!(obj is EventRecurrence))
                return false;
            EventRecurrence target = obj as EventRecurrence;
            bool exceptionsEqual;
            if (this.Exceptions == null || this.Exceptions.Count == 0)
            {
                exceptionsEqual = target.Exceptions == null || target.Exceptions.Count == 0;
            }
            else
            {
                exceptionsEqual = (target.Exceptions != null) && (target.Exceptions.Count != 0) 
                    && this.Exceptions.SequenceEqual(target.Exceptions, new RecurrenceExceptionComparer());
            }
            return
                this.Count == target.Count &&
                this.DayOfMonth == target.DayOfMonth &&
                this.DayOfWeekMask == target.DayOfWeekMask &&
                this.EndDate == target.EndDate &&
                this.EndMethod == target.EndMethod &&
                this.Frequency == target.Frequency &&
                this.Interval == target.Interval &&
                this.Month == target.Month &&
                this.StartDate == target.StartDate &&
                this.WeekInterval == target.WeekInterval &&
                exceptionsEqual;
        }

        private DayOfWeek GetDayOfWeekMask(string value)
        {
            DayOfWeek res = 0;
            var regex = new Regex("\\d?(\\w{2},?)", RegexOptions.Compiled);
            value = regex.Replace(value, "$1");
            foreach (var dayOfWeekAbbr in value.Split(','))
                if (dayOfWeekAbbr == "MO")
                    res |= DayOfWeek.Monday;
                else if (dayOfWeekAbbr == "TU")
                    res |= DayOfWeek.Tuesday;
                else if (dayOfWeekAbbr == "WE")
                    res |= DayOfWeek.Wednesday;
                else if (dayOfWeekAbbr == "TH")
                    res |= DayOfWeek.Thursday;
                else if (dayOfWeekAbbr == "FR")
                    res |= DayOfWeek.Friday;
                else if (dayOfWeekAbbr == "SA")
                    res |= DayOfWeek.Saturday;
                else if (dayOfWeekAbbr == "SU")
                    res |= DayOfWeek.Sunday;
            return res;
        }

        public override int GetHashCode()
        {
            return
                this.GetType().GetProperties().Aggregate(
                    0,
                    (hash, prop) =>
                        hash ^ (int)prop.DeclaringType.InvokeMember(
                            "GetHashCode",
                            System.Reflection.BindingFlags.InvokeMethod,
                            null,
                            this.GetType().InvokeMember(
                                prop.Name,
                                System.Reflection.BindingFlags.GetProperty,
                                null,
                                this,
                                null),
                            null));
        }
    }
}
