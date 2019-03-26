using Google.Apis.Calendar.v3.Data;
using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace R.GoogleOutlookSync
{
    class EventComparer
    {
        internal static bool Equals(Event googleItem, Outlook.AppointmentItem outlookItem)
        {
            var privacyMode = (int)Utilities.GetRegistryValue(VSTO.Properties.Settings.Default.Privacy) == 1;
            var attendeesEqual = privacyMode ? true : AttendeeComparer.Equals(googleItem.Attendees, outlookItem.Recipients);
            var bodiesEqual = privacyMode ? true : googleItem.Description == outlookItem.Body;
            var locationsEqual = privacyMode ? true : LocationIsEqual(googleItem, outlookItem);
            return
                googleItem.Summary == outlookItem.Subject &&
                bodiesEqual &&
                attendeesEqual &&
                locationsEqual &&
                ReminderIsEqual(googleItem, outlookItem);
        }

        internal static bool Equals(CalendarEvent x, CalendarEvent y)
        {
            var privacyMode = (int)Utilities.GetRegistryValue(VSTO.Properties.Settings.Default.Privacy) == 1;
            var attendeesEqual = privacyMode ? true : x.Attendees.SequenceEqual(y.Attendees, new AttendeeComparer());
            var bodiesEqual = privacyMode ? true : StringIsEqual(x.Body, y.Body);
            var locationsEqual = privacyMode ? true : StringIsEqual(x.Location, y.Location);

            return
                StringIsEqual(x.Subject, y.Subject) &&
                bodiesEqual &&
                attendeesEqual &&
                locationsEqual &&
                ReminderIsEqual(x, y);
        }

        private static bool StringIsEqual(string x, string y)
        {
            return
                (String.IsNullOrEmpty(x) && string.IsNullOrEmpty(y)) ||
                (x == y);
        }

        [FieldComparer(Field.Location)]
        private static bool LocationIsEqual(Event googleItem, Outlook.AppointmentItem outlookItem)
        {
            return
                (String.IsNullOrEmpty(googleItem.Location) && String.IsNullOrEmpty(outlookItem.Location)) ||
                (googleItem.Location == outlookItem.Location);
        }

        [FieldComparer(Field.Reminder)]
        private static bool ReminderIsEqual(Event googleItem, Outlook.AppointmentItem outlookItem)
        {
            if ((googleItem.Reminders == null) && !outlookItem.ReminderSet)
                return true;
            if (googleItem.Reminders != null)
                return googleItem.Reminders.Overrides.First().Minutes == outlookItem.ReminderMinutesBeforeStart;
            return false;
        }

        private static bool ReminderIsEqual(CalendarEvent x, CalendarEvent y)
        {
            return
                (x.ReminderSet == y.ReminderSet) &&
                (!x.ReminderSet || x.ReminderMinutes == y.ReminderMinutes);
        }
    }
}
