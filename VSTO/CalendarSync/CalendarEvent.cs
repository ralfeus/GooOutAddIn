using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3.Data;

namespace R.GoogleOutlookSync
{
    class CalendarEvent
    {
        internal EventSchedule Schedule { get; set; }
        internal string Subject { get; set; }
        internal string Body { get; set; }
        internal bool ReminderSet { get; set; }
        internal int ReminderMinutes { get; set; }
        internal List<Attendee> Attendees { get; set; }
        internal string Location { get; set; }

        internal CalendarEvent(Event googleItem, IList<EventReminder> googleDefaultReminders)
        {
            this.Subject = googleItem.Summary;
            this.Body = googleItem.Description;
            this.ReminderSet = googleItem.Reminders != null;
            if (this.ReminderSet) {
                if (googleItem.Reminders.UseDefault.Value) {
                    this.ReminderMinutes = googleDefaultReminders.FirstOrDefault().Minutes.Value;
                } else {
                    this.ReminderMinutes = googleItem.Reminders.Overrides.FirstOrDefault().Minutes.Value;
                }
            } else {
                this.ReminderMinutes = 0;
            }
            if (googleItem.Attendees != null) {
                this.Attendees = new List<Attendee>(googleItem.Attendees.Count - 1);
                foreach (EventAttendee participant in googleItem.Attendees)
                    if (participant.Organizer.Value)
                        this.Attendees.Add(new Attendee(participant.Email, GoogleUtilities.GetAttendeeType(participant), participant.ResponseStatus));
            }
            this.Location = googleItem.Location;
            this.Schedule = new EventSchedule(googleItem);
        }

        internal CalendarEvent(AppointmentItem outlookItem)
        {
            this.Subject = outlookItem.Subject;
            this.Body = outlookItem.Body;
            this.ReminderSet = outlookItem.ReminderSet;
            this.ReminderMinutes = outlookItem.ReminderMinutesBeforeStart;
            this.Attendees = new List<Attendee>(outlookItem.Recipients.Count);
            foreach (Recipient recipient in outlookItem.Recipients)
                this.Attendees.Add(new Attendee(recipient.Address, recipient.Type, recipient.MeetingResponseStatus));
            this.Location = outlookItem.Location;
            this.Schedule = new EventSchedule(outlookItem);
        }

        public override bool Equals(object obj)
        {
            if (object.ReferenceEquals(this, obj))
                return true;
            CalendarEvent target = (CalendarEvent)obj;
            return
                this.Attendees.SequenceEqual(target.Attendees, new AttendeeComparer()) &&
                this.Body == target.Body &&
                this.Location == target.Location &&
                this.ReminderSet == target.ReminderSet &&
                this.ReminderMinutes == target.ReminderMinutes &&
                this.Subject == target.Subject &&
                this.Schedule.Equals(target.Schedule);
        }

        public override int GetHashCode()
        {
            return
                this.Attendees.GetHashCode() ^
                this.Body.GetHashCode() ^
                this.Location.GetHashCode() ^
                this.ReminderMinutes.GetHashCode() ^
                this.ReminderSet.GetHashCode() ^
                this.Schedule.GetHashCode() ^
                this.Subject.GetHashCode();
        }

        internal void ToGoogle(Event googleItem)
        {
            this.Schedule.ToGoogle(googleItem);

            if (this.ReminderSet)
            {
                googleItem.Reminders = new Event.RemindersData();
                googleItem.Reminders.Overrides.Add(new EventReminder
                {
                    Minutes = this.ReminderMinutes
                });
            }
        }

        internal Event ToGoogle()
        {
            Event googleItem = new Event()
            {
                Summary = this.Subject,
                Description = this.Body,
                Location = this.Location
            };
            this.ToGoogle(googleItem);
            return googleItem;
        }
    }
}