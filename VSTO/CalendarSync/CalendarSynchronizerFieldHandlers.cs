using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Discovery;

namespace R.GoogleOutlookSync
{
    internal partial class CalendarSynchronizer
    {
        /// <summary>
        /// Compares AllDayEvent attribute for items
        /// </summary>
        /// <param name="googleItem"></param>
        /// <param name="outlookItem"></param>
        /// <returns></returns>
        //[FieldComparer(Field.AllDayEvent)]
        private bool CompareAllDayEvent(Event googleItem, object outlookItem)
        {
            var result = (googleItem.Start.Date != "") == ((Outlook.AppointmentItem)outlookItem).AllDayEvent;
            return result;
        }

        //private bool CompareRecurrence(Event googleItem, Outlook.AppointmentItem outlookItem)
        //{
        //    EventRecurrence googleRec = null;
        //    EventRecurrence outlookRec = null;
        //    try
        //    {
        //        if (googleItem.Recurrence != null)
        //            googleRec = new EventRecurrence(googleItem.Recurrence);
        //    }
        //    catch (ArgumentNullException) { }
        //    try
        //    {
        //        if ((Outlook.AppointmentItem)outlookItem.IsRecurring)
        //            outlookRec = new EventRecurrence((Outlook.AppointmentItem)outlookItem.GetRecurrencePattern());
        //    }
        //    catch (ArgumentNullException) { }
        //    return googleRec == outlookRec;
        //}

        /// Attachments are not accessible yet by Google API
        [FieldComparer(Field.Attachments)]
        private bool CompareAttachments(Event googleItem, object outlookItem)
        {
            return true;
        }

        [FieldComparer(Field.Attendees)]
        private bool CompareAttendees(Event googleItem, object outlookItem)
        {
            return AttendeeComparer.Equals(googleItem.Attendees, ((Outlook.AppointmentItem)outlookItem).Recipients); // ((vbMAPI_AppointmentItem)vbMAPI_Init.NewOutlookWrapper((Outlook.AppointmentItem)outlookItem)).Recipients);
        }

        [FieldComparer(Field.Description)]
        private bool CompareDescription(Event googleItem, object outlookItem)
        {
            //var outlookItemBody = ((vbMAPI_AppointmentItem)vbMAPI_Init.NewOutlookWrapper((Outlook.AppointmentItem)outlookItem)).Body.Value;
            return
                (String.IsNullOrEmpty(googleItem.Description) && String.IsNullOrEmpty(((Outlook.AppointmentItem)outlookItem).Body)) ||
                (googleItem.Description == ((Outlook.AppointmentItem)outlookItem).Body);
        }

        [FieldComparer(Field.Location)]
        private bool CompareLocation(Event googleItem, object outlookItem)
        {
            return
                (String.IsNullOrEmpty(googleItem.Location) && String.IsNullOrEmpty(((Outlook.AppointmentItem)outlookItem).Location)) ||
                (googleItem.Location == ((Outlook.AppointmentItem)outlookItem).Location);
        }

        //[FieldComparer(Field.Reminder)]
        private bool CompareReminder(Event googleItem, object outlookItem)
        {
            var googleReminders = (!googleItem.Reminders.UseDefault.HasValue || googleItem.Reminders.UseDefault.Value)
                ? this._defaultReminders
                : googleItem.Reminders.Overrides;
            if (googleReminders == null) 
            {
                return !((Outlook.AppointmentItem)outlookItem).ReminderSet;
            }
            else
            {
                // if Google appointment has more than 1 reminder, we ignore reminders as Outlook can't have more than 1
                if (googleReminders.Count > 1)
                {
                    Logger.Log(String.Format(VSTO.Properties.Resources.Info_MultipleReminders, googleItem.Summary, googleReminders.Count), EventType.Information);
                    return true;
                }
                return 
                    ((Outlook.AppointmentItem)outlookItem).ReminderSet  && 
                    (googleReminders.First().Minutes == ((Outlook.AppointmentItem)outlookItem).ReminderMinutesBeforeStart);
            }
        }

        [FieldComparer(Field.ShowTimeAs)]
        [Public]
        private bool CompareShowTimeAs(Event googleItem, object outlookItem)
        {
            return (googleItem.Transparency ?? "opaque") == ConvertTo.GoogleAvailability(((Outlook.AppointmentItem)outlookItem).BusyStatus);
        }
        
        [FieldComparer(Field.Subject)]
        [Public]
        private bool CompareSubjects(Event googleItem, object outlookItem)
        {
            return googleItem.Summary.Equals(((Outlook.AppointmentItem)outlookItem).Subject);
        }

        [FieldComparer(Field.Time)]
        [Public]
        private bool CompareTime(Event googleItem, object outlookItem)
        {
            var instancesRequest = CalendarService.Events.Instances(_googleCalendar.Id, googleItem.Id);
            instancesRequest.ShowDeleted = true;
            var googleExceptions = instancesRequest.Execute().Items.Where(e => e.Status == "cancelled").ToList();
            var googleEventSchedule = new EventSchedule(googleItem, this._googleCalendar, this.CalendarService /*googleExceptions*/);
            var outlookEventSchedule = new EventSchedule((Outlook.AppointmentItem)outlookItem);
            
            return
                googleEventSchedule == outlookEventSchedule &&
                this.CompareReminder(googleItem, outlookItem);
        }

        [FieldComparer(Field.Visibility)]
        [Public]
        private bool CompareVisibility(Event googleItem, object outlookItem)
        {
            return (googleItem.Visibility ?? "default" ) == ConvertTo.GoogleVisibility(((Outlook.AppointmentItem)outlookItem).Sensitivity);
        }

        //[FieldSetter(Field.AllDayEvent)]
        //private void SetAllDayEvent(Event googleItem, object outlookItem, Target target)
        //{
        //    if (target == Target.Google)
        //    {
        //        if (googleItem.Times.Count == 0)
        //            googleItem.Times.Add(new When());
        //        googleItem.Times[0].AllDay = ((Outlook.AppointmentItem)outlookItem).AllDayEvent;
        //    }
        //    else
        //        ((Outlook.AppointmentItem)outlookItem).AllDayEvent = googleItem.Times[0].AllDay;
        //}

        /// Attachments are not accessible yet by Google API
        [FieldSetter(Field.Attachments)]
        private void SetAttachments(Event googleItem, object outlookItem, Target target)
        {
        }

        [FieldSetter(Field.Attendees)]
        private void SetAttendees(Event googleItem, object outlookItem, Target target)
        {
            if (target == Target.Google)
            {
                if (googleItem.Attendees == null)
                {
                    googleItem.Attendees = new List<EventAttendee>();
                }
                foreach (Outlook.Recipient outlookRecipient in ((Outlook.AppointmentItem)outlookItem).Recipients)
                {
                    try 
                    {
#if DEBUG
                        Logger.Log(string.Format("Adding attendee {0} to Google event", outlookRecipient.Address), EventType.Debug);
#endif
                        /// Organizator is omited
                        if (((Outlook.OlMeetingRecipientType)outlookRecipient.Type == Outlook.OlMeetingRecipientType.olOrganizer) ||
                            (outlookRecipient.Address == null) ||
                            !Utilities.SMTPAddressPattern.IsMatch(outlookRecipient.Address))
                        {
#if DEBUG
                            Logger.Log(string.Format("{0} is a meeting organizer or it's not valid SMTP address, so it won't be added", outlookRecipient.Address), 
                                EventType.Debug);
#endif
                            continue;
                        }
                        var googleRecipient = googleItem.Attendees.FirstOrDefault(recipient => recipient.Email == outlookRecipient.Address);
                        if (googleRecipient == null)
                        {
                            googleRecipient = new EventAttendee
                            {
                                Email = outlookRecipient.Address
                            };
                            googleItem.Attendees.Add(googleRecipient);
                        }
                        else if (AttendeeComparer.Equals(googleRecipient, outlookRecipient))
                        {
                            continue;
                        }
                        googleRecipient = ConvertTo.GoogleRecipientType(googleRecipient, (Outlook.OlMeetingRecipientType)outlookRecipient.Type);
                        googleRecipient.ResponseStatus = ConvertTo.GoogleResponseStatus(outlookRecipient.MeetingResponseStatus);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(outlookRecipient);
                    }
                }
            }
            else
            {
                if (googleItem.Attendees != null)
                {
                    foreach (var googleRecipient in googleItem.Attendees)
                    {
                        /// Organizator and not valid SMTP addresses are omited
                        if (googleRecipient.Organizer.HasValue && googleRecipient.Organizer.Value ||
                            !Utilities.SMTPAddressPattern.IsMatch(googleRecipient.Email))
                            continue;
                        var matchIsFound = false;
                        foreach (Outlook.Recipient outlookRecipient in ((Outlook.AppointmentItem)outlookItem).Recipients)
                        {
                            if (googleRecipient.Email == outlookRecipient.Address)
                            {
                                matchIsFound = true;
                                Marshal.ReleaseComObject(outlookRecipient);
                                break;
                            }
                            Marshal.ReleaseComObject(outlookRecipient);
                        }
                        if (!matchIsFound)
                        {
                            Outlook.Recipient outlookRecipient = ((Outlook.AppointmentItem)outlookItem).Recipients.Add(googleRecipient.Email);
                            outlookRecipient.Type = (int)ConvertTo.OutlookRecipientType(googleRecipient);
                            Marshal.ReleaseComObject(outlookRecipient);
                        }
                    }
                }
            }
        }

        [FieldSetter(Field.Description)]
        private void SetDescription(Event googleItem, object outlookItem, Target target)
        {
            if (target == Target.Google)
                googleItem.Description = ((Outlook.AppointmentItem)outlookItem).Body;
            else
                ((Outlook.AppointmentItem)outlookItem).Body = googleItem.Description;
        }

        [FieldSetter(Field.Location)]
        private void SetLocation(Event googleItem, object outlookItem, Target target)
        {
            if (target == Target.Google)
            {
                    googleItem.Location = ((Outlook.AppointmentItem)outlookItem).Location;
            }
            else
                ((Outlook.AppointmentItem)outlookItem).Location = googleItem.Location;
        }

        /// <summary>
        /// Sets recurrence in direction Google -> Outlook
        /// </summary>
        /// <param name="googleItem">Source Google event</param>
        /// <param name="outlookItem">Target Outlook event</param>
        private void SetRecurrence(Event googleItem, Outlook.AppointmentItem outlookItem)
        {
            if (googleItem.Recurrence != null)
            {
                Logger.Log(string.Join(@"\n", googleItem.Recurrence), EventType.Debug);
                new EventSchedule(googleItem, this._googleCalendar, this.CalendarService).GetOutlookRecurrence(outlookItem);
                this.SetRecurrenceExceptions(googleItem, outlookItem);
            }
            else
                outlookItem.ClearRecurrencePattern();
        }
        
        /// <summary>
        /// Sets recurrence in direction Outlook -> Google
        /// </summary>
        /// <param name="outlookItem">Source Outlook event</param>
        /// <param name="googleItem">Target Google event</param>
        private void SetRecurrence(Outlook.AppointmentItem outlookItem, Event googleItem)
        {
            if (outlookItem.IsRecurring)
            {
                /// Times property doesn't contain anything for recurrent events
                var outlookItemSchedule = new EventSchedule(outlookItem);
                googleItem.Recurrence = outlookItemSchedule.GetGoogleRecurrence();
                /// Recurrence exceptions can be set for Google item only after the item is create or updated
                /// Only in this case all recurrence instances are created
                //this.SetRecurrenceExceptions((outlookItem), googleItem);
            }
            else
            {
                googleItem.Recurrence = null;
            }
        }

        //[FieldSetter(Field.Reminder)]
        private void SetReminder(Event googleItem, Outlook.AppointmentItem outlookItem, Target target)
        {
            if (target == Target.Google)
            {
                if (outlookItem.ReminderSet)
                {
                    if (outlookItem.ReminderOverrideDefault)
                    {
                        var reminder = new EventReminder
                        {
                            Minutes = outlookItem.ReminderMinutesBeforeStart,
                            Method = "popup"
                        };
                        googleItem.Reminders.Overrides = new[] { reminder };
                    } else
                    {
                        googleItem.Reminders.Overrides = null;
                    }
                    googleItem.Reminders.UseDefault = !outlookItem.ReminderOverrideDefault;
                }
                else
                {
                    googleItem.Reminders = null;
                }
            }
            else if (googleItem.Reminders == null)
            {
                ((Outlook.AppointmentItem)outlookItem).ReminderSet = false;
            }
            else
            {
                if (googleItem.Reminders.UseDefault.GetValueOrDefault())
                {
                    if (this._defaultReminders.Any())
                    {
                        ((Outlook.AppointmentItem)outlookItem).ReminderSet = true;
                        ((Outlook.AppointmentItem)outlookItem).ReminderOverrideDefault = true;
                        ((Outlook.AppointmentItem)outlookItem).ReminderMinutesBeforeStart = this._defaultReminders.First().Minutes.GetValueOrDefault();
                    } else
                    {
                        ((Outlook.AppointmentItem)outlookItem).ReminderSet = false;
                    }
                } else
                {
                    if (googleItem.Reminders.Overrides != null && googleItem.Reminders.Overrides.Count > 0)
                    {
                        ((Outlook.AppointmentItem)outlookItem).ReminderSet = true;
                        ((Outlook.AppointmentItem)outlookItem).ReminderOverrideDefault = true;
                        ((Outlook.AppointmentItem)outlookItem).ReminderMinutesBeforeStart = googleItem.Reminders.Overrides.First().Minutes.GetValueOrDefault();
                    } else
                    {
                        ((Outlook.AppointmentItem)outlookItem).ReminderSet = false;
                    }
                }
            }
        }

        /// <summary>
        /// Sets recurrence exceptions in direction Google -> Outlook
        /// </summary>
        /// <param name="googleItem">Source Google event</param>
        /// <param name="outlookItem">Target Outlook event</param>
        private void SetRecurrenceExceptions(Event googleItem, object outlookItem)
        {
            //if (!this._googleExceptions.ContainsKey(googleItem.Id))
            //    return;
            ((Outlook.AppointmentItem)outlookItem).Save();
            Outlook.RecurrencePattern outlookRecurrence = ((Outlook.AppointmentItem)outlookItem).GetRecurrencePattern();
            try
            {
                if (this._googleExceptions.ContainsKey(googleItem.Id))
                {
                    foreach (Event googleException in this._googleExceptions[googleItem.Id])
                    {
                        /// If the exception is already in Outlook event this one is omited
                        //if (RecurrenceExceptionComparer.Contains(outlookRecurrence.Exceptions, googleException))
                        //    continue;
                        /// Get occurence of the recurrence and modify it. Thus new exception is created
                        Outlook.AppointmentItem outlookExceptionItem = outlookRecurrence.GetOccurrence(googleException.OriginalStartTime.DateTime.Value);
                        try
                        {
                            if (googleException.Status == "cancelled")
                            {
                                outlookExceptionItem.Delete();
                            }
                            else
                            {
                                this.SetSubject(googleException, outlookExceptionItem, Target.Outlook);
                                this.SetDescription(googleException, outlookExceptionItem, Target.Outlook);
                                this.SetLocation(googleException, outlookExceptionItem, Target.Outlook);
                                this.SetTime(googleException, outlookExceptionItem, Target.Outlook);
                                outlookExceptionItem.Save();
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(outlookExceptionItem);
                        }
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(outlookRecurrence);
            }
        }

        /// <summary>
        /// Sets recurrence exceptions in direction Outlook -> Google
        /// </summary>
        /// <param name="outlookItem">Source Outlook event</param>
        /// <param name="googleItem">Target Google event</param>
        private void SetRecurrenceExceptions(Outlook.AppointmentItem outlookItem, Event googleItem)
        {
            if (!outlookItem.IsRecurring)
            {
                return;
            }
            Outlook.RecurrencePattern outlookRecurrence = outlookItem.GetRecurrencePattern();
            var instancesRequest = this.CalendarService.Events.Instances(this._googleCalendar.Id, googleItem.Id);
            instancesRequest.ShowDeleted = true;
            var instances = instancesRequest.Execute().Items;
            var exceptions = instances.Where(i => i.Status == "cancelled" || i.OriginalStartTime.DateTime != i.Start.DateTime);
            try
            {
                foreach (Outlook.Exception outlookException in outlookRecurrence.Exceptions)
                {
                    try {
                        /// If the exception is already in Google event this one is omited
                        if (RecurrenceExceptionComparer.Contains(exceptions, outlookException))
                            continue;
                        //googleItem.RecurrenceException = new Google.GData.Extensions.RecurrenceException();
                        Event googleException;
                        if (outlookException.Deleted) {
                            googleException = instances.First(e =>
                            (e.OriginalStartTime.DateTime ?? DateTime.Parse(e.OriginalStartTime.Date)).Date == outlookException.OriginalDate);
                            googleException.Status = "cancelled";
                        } else {
                            googleException = instances.First(e =>
                            (e.OriginalStartTime.DateTime ?? DateTime.Parse(e.OriginalStartTime.Date)) == outlookException.OriginalDate);
                            googleException.Status = "confirmed";
                            this.SetSubject(googleException, outlookException.AppointmentItem, Target.Google);
                            this.SetDescription(googleException, outlookException.AppointmentItem, Target.Google);
                            this.SetLocation(googleException, outlookException.AppointmentItem, Target.Google);
                            this.SetTime(googleException, outlookException.AppointmentItem, Target.Google);
                        }
                        /// Add exception to update batch
                        this._googleBatchRequest.Queue<Event>(this.CalendarService.Events.Update(
                            googleException, this._googleCalendar.Id, googleException.Id), BatchCallback);
                    } catch (COMException exc) {
                        if (exc.HResult == -1802485755) {
                            Logger.Log(string.Format(
                                "The exception of '{0}' with original start time {1} couldn't be read. Ignoring this exception",
                                outlookItem.Subject, outlookException.OriginalDate), EventType.Warning);
                        }
                    } finally {
                        Marshal.ReleaseComObject(outlookException);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(outlookRecurrence);
            }
        }

        /// <summary>
        /// Sets event transparency
        /// </summary>
        /// <param name="googleItem">Google event</param>
        /// <param name="outlookItem">Outlook event</param>        
        [FieldSetter(Field.ShowTimeAs)]
        [Public]
        private void SetShowTimeAs(Event googleItem, object outlookItem, Target target)
        {
            if (target == Target.Google)
                googleItem.Transparency = ConvertTo.GoogleAvailability(((Outlook.AppointmentItem)outlookItem).BusyStatus);
            else
                ((Outlook.AppointmentItem)outlookItem).BusyStatus = ConvertTo.OutlookAvailability(googleItem.Transparency);
        }

        /// <summary>
        /// Sets event schedule
        /// </summary>
        /// <param name="googleItem">Google event</param>
        /// <param name="outlookItem">Outlook event</param>        
        [FieldSetter(Field.Time)]
        [Public]
        private void SetTime(Event googleItem, object item, Target target)
        {
            var outlookItem = (Outlook.AppointmentItem)item;
            /// in order to prevent situation when start is later then end
            /// we should process both times simultaniously and check what time should be changed first
            if (target == Target.Google)
            {
                //var schedule = new EventSchedule(((Outlook.AppointmentItem)outlookItem));
                if (outlookItem.AllDayEvent)
                {
                    googleItem.Start.Date = outlookItem.Start.ToString("yyyy-MM-dd");
                    googleItem.End.Date = outlookItem.End.ToString("yyyy-MM-dd");
                }
                else
                {
                    googleItem.Start.DateTime = outlookItem.Start;
                    googleItem.Start.TimeZone = TimeZones.GetTZ(outlookItem.StartTimeZone.ID);
                    googleItem.End.DateTime = outlookItem.End;
                    googleItem.End.TimeZone = TimeZones.GetTZ(outlookItem.EndTimeZone.ID);
                }
                if (outlookItem.IsRecurring && (outlookItem.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster))
                {
                    this.SetRecurrence(outlookItem, googleItem);
                }
            }
            else
            {
                if (googleItem.Recurrence == null)
                {
                    var schedule = new EventSchedule(googleItem);
                    outlookItem.Start = schedule.StartTime;
                    outlookItem.End = schedule.EndTime;
                    outlookItem.AllDayEvent = schedule.AllDayEvent;
                }
                else
                    this.SetRecurrence(googleItem, outlookItem);
            }
            this.SetReminder(googleItem, outlookItem, target);
        }

        /// <summary>
        /// Sets event subject
        /// </summary>
        /// <param name="googleItem">Google event</param>
        /// <param name="outlookItem">Outlook event</param>        
        [FieldSetter(Field.Subject)]
        [Public]
        private void SetSubject(Event googleItem, object outlookItem, Target target)
        {
            if (target == Target.Google)
            {
                googleItem.Summary = ((Outlook.AppointmentItem)outlookItem).Subject;
            }
            else
            {
                ((Outlook.AppointmentItem)outlookItem).Subject = googleItem.Summary;
            }
        }

        /// <summary>
        /// Sets event visibility
        /// </summary>
        /// <param name="googleItem">Google event</param>
        /// <param name="outlookItem">Outlook event</param>        
        [FieldSetter(Field.Visibility)]
        [Public]
        private void SetVisibility(Event googleItem, object outlookItem, Target target)
        {
            if (target == Target.Google)
                googleItem.Visibility = ConvertTo.GoogleVisibility(((Outlook.AppointmentItem)outlookItem).Sensitivity);
            else 
                ((Outlook.AppointmentItem)outlookItem).Sensitivity = ConvertTo.OutlookVisibility(googleItem.Visibility);
        }
    }
}