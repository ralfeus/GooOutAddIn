using Google.Apis.Calendar.v3.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace R.GoogleOutlookSync
{
    class RecurrenceException
    {
        private EventRecurrence _parentRecurrence;

        internal bool Deleted { get; set; }
        internal CalendarEvent ModifiedEvent { get; set; }
        internal DateTime OriginalDate { get; set; }

        internal RecurrenceException(EventRecurrence parentRecurrence, DateTime originalDate)
        {
            this._parentRecurrence = parentRecurrence;
            this.OriginalDate = originalDate;
        }

        internal RecurrenceException(Event googleException, IList<EventReminder> googleDefaultReminders, EventRecurrence parentRecurrence = null)
        {
            this._parentRecurrence = parentRecurrence ?? null;
            this.OriginalDate = googleException.OriginalStartTime.DateTime.HasValue
                ? googleException.OriginalStartTime.DateTime.Value
                : DateTime.Parse(googleException.OriginalStartTime.Date);
            this.Deleted = googleException.Status == "cancelled";
            if (!this.Deleted)
                this.ModifiedEvent = new CalendarEvent(googleException, googleDefaultReminders);
        }

        internal RecurrenceException(Outlook.Exception outlookException, EventRecurrence parentRecurrence = null)
        {
            this._parentRecurrence = parentRecurrence != null ? parentRecurrence : new EventRecurrence(outlookException.Parent as Outlook.RecurrencePattern);
            this.OriginalDate = outlookException.OriginalDate;
            this.Deleted = outlookException.Deleted;
            if (!this.Deleted)
                this.ModifiedEvent = new CalendarEvent(outlookException.AppointmentItem);
        }
    }
}
