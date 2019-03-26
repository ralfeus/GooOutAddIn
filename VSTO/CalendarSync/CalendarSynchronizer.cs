using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Calendar.v3;
using Google.Apis.Requests;
using System.Threading;
using System.Net.Http;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using Google;

namespace R.GoogleOutlookSync
{
    internal partial class CalendarSynchronizer : ItemTypeSynchronizer
    {
        /// <summary>
        /// Wrapper around common Google service for CalendarService
        /// </summary>
        private CalendarService CalendarService { get => (CalendarService)this._googleService; }
        /// <summary>
        /// Represents calendar to synchronize with
        /// </summary>
        Calendar _googleCalendar;
        private IList<EventReminder> _defaultReminders;
        string _googleCalendarName;

        /// <summary>
        /// Contains exceptions from recurrent events
        /// </summary>
        Dictionary<string, List<Event>> _googleExceptions;
        //new IEnumerable<FieldHandlers<Outlook.AppointmentItem>> _fieldHandlers;
        /// <summary>
        /// Contains 0:00 of next date after defined interval. Next date is not inclusive
        /// </summary>
        DateTime _rangeEnd;
        DateTime _rangeStart;

        internal CalendarSynchronizer(
            UserCredential googleCredential,
            Outlook.Application outlookApplication,
            string calendarFolderToSyncID)
        {
            this._googleService = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = googleCredential
            });
            this.outlookApplication = outlookApplication;

            this._outlookFolderToSyncID = calendarFolderToSyncID;
            this._googleExceptions = new Dictionary<string, List<Event>>();
            //this._googleBatchRequest = new BatchRequest(this.CalendarService);
            
            this.LoadSettings();
            var calendars = this.CalendarService.CalendarList.List().Execute().Items;
            // first load calendars synchronization settings to have Google calendar name to synchronize
            // in case no settings are loaded first found calendar will be used
            foreach (var cal in calendars)
            {
                Logger.Log("Found calendar: " + cal.Summary, EventType.Information);
            }
            var calendar = calendars.First(
                cal => (cal.Summary == this._googleCalendarName) || String.IsNullOrEmpty(this._googleCalendarName));
            this._googleCalendar = this.CalendarService.Calendars.Get(calendar.Id).Execute();
            this._defaultReminders = calendar.DefaultReminders;
            /// Prepare list of all available field handlers
            this.InitFieldHandlers();
        }

        /// <summary>
        /// Get last modification time for Google event. Check also time of exceptions modification 
        /// </summary>
        /// <param name="googleItem"></param>
        /// <returns></returns>
        protected override DateTime GetLastModificationTime(Event googleItem)
        {
            if (this._googleExceptions.ContainsKey(((Event)googleItem).Id))
                return
                    new DateTime(Math.Max(
                        GoogleUtilities.GetLastModificationTime(googleItem).Ticks,
                        this._googleExceptions[((Event)googleItem).Id].Max(exception => exception.Updated.Value.Ticks)));
            else
                return GoogleUtilities.GetLastModificationTime(googleItem);
        }

        protected override ComparisonResult Compare(Event googleItem, object item)
        {
            /// If IDs are same items are same. Then we can check whether they are identical
            /// If IDs are different items are different
            /// Aggregate will invoke each method in _fieldComparers list. Boolean AND will ensure all comparers return true
            //try
            //{
            Outlook.AppointmentItem appointment = (Outlook.AppointmentItem)item;
            if (this.CompareIDs((Event)googleItem, appointment))
            {
                var fieldsAreSame =
                    this._fieldHandlers.Aggregate(
                        true,
                        (res, handler) =>
                        {
                            Logger.Log(string.Format("Comparing '{0}' of {1} at {2}", handler.Comparer.Method.Name, appointment.Subject, appointment.Start), EventType.Debug);
                            var result = handler.Comparer(googleItem, appointment);
                            Logger.Log(string.Format("\t{0}", result), EventType.Debug);
                            return res && result;
                        }
                    );
                //var fieldsAreSame = true;
                //foreach (var fieldHandler in this._fieldHandlers)
                //    fieldsAreSame &= (bool)fieldHandler.Comparer.Invoke(this, new object[] { googleItem, outlookItem });
                if (fieldsAreSame)
                    return ComparisonResult.Identical;
                else
                    return ComparisonResult.SameButChanged;
            }
            else
                return ComparisonResult.Different;
            //}
            //catch (TargetInvocationException exc)
            //{
            //    throw exc.InnerException;
            //}
        }

        protected override bool IsItemValid(Event googleItem)
        {
            /// We try to create EventSchedule of item. If due to any reason (Exception) it's impossible
            /// we treat item as invalid
            try
            {
                new EventSchedule((Event)googleItem);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        protected override bool IsItemValid(object outlookItem)
        {
            /// So far we treat all Outlook items as valid
            return true;
        }

        protected override void LoadGoogleItems()
        {
            Logger.Log("Loading Google calendar items...", EventType.Information);
            var succeeded = false;
            var attemptsLeft = 3;
            Exception lastError = null;
            do
            {
                try
                {
                    var query = this.CalendarService.Events.List(this._googleCalendar.Id);
                    query.TimeMin = this._rangeStart;
                    query.TimeMax = this._rangeEnd;
                    query.ShowDeleted = true;
                    var events = query.Execute().Items;
                    this.GoogleItems = events.Where(e => e.Status != "cancelled").ToList();
                    foreach (var e in events.Where(e => !string.IsNullOrEmpty(e.RecurringEventId)))
                    {
                        if (!this._googleExceptions.ContainsKey(e.RecurringEventId))
                        {
                            this._googleExceptions.Add(e.RecurringEventId, new List<Event>());
                        }
                        this._googleExceptions[e.RecurringEventId].Add(e);
                    }
                    succeeded = true;
                }
                catch (Exception exc)
                {
                    ErrorHandler.Handle(exc);
                    lastError = exc;
                    --attemptsLeft;
                }
            } while (!succeeded && attemptsLeft > 0);
        }

        protected override void LoadOutlookItems()
        {
            Logger.Log("Loading Outlook calendar items...", EventType.Information);
            var calendarItems = this.GetOutlookItems(this._outlookFolderToSyncID);
            Logger.Log("Sorting Outlook calendar items", EventType.Debug);
            calendarItems.Sort("Start");
            //calendarItems.IncludeRecurrences = true;
            //var query = "NOT(\"urn:schemas:calendar:dtstart\" IS NULL)";
            var query = String.Format("@SQL=NOT(\"urn:schemas:calendar:dtstart\" IS NULL)");
            Logger.Log("Filtering Outlook calendar items", EventType.Debug);
            Outlook.AppointmentItem item = calendarItems.Find(query) as Outlook.AppointmentItem;

            while (item != null)
            {
                Outlook.AppointmentItem tmpEvent;
                // any MeetingItem is actually Appointment with special status. So we get correspondent appointmen
                if (item is Outlook.MeetingItem)
                    tmpEvent = ((Outlook.MeetingItem)item).GetAssociatedAppointment(false);
                else if (item is Outlook.AppointmentItem)
                    tmpEvent = item;
                // it's also possible non-calendar item can be placed in the Calendar folder. We'll just ignore all such items
                else
                    continue;
                if ((tmpEvent.Start <= this._rangeEnd) && (tmpEvent.End > this._rangeStart))
                {
                    this.OutlookItems.Add(item);
                }
                else if (tmpEvent.IsRecurring)
                {
                    var pattern = tmpEvent.GetRecurrencePattern();
                    if ((pattern.PatternStartDate <= this._rangeEnd) && (pattern.PatternEndDate >= this._rangeStart))
                    {
                        this.OutlookItems.Add(item);
                    }
                }
                //// release appoinment item since it's COM object 
                //Marshal.ReleaseComObject(item);
                //Marshal.ReleaseComObject(tmpEvent);
                item = (Outlook.AppointmentItem)calendarItems.FindNext();
            }
            Marshal.ReleaseComObject(calendarItems);
        }

        private void LoadSettings()
        {
            // if start and end synchronization range can't be read from the registry default values will be used
            var daysBefore = 0;
            var daysAfter = 0;
            try
            {
                var regKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(VSTO.Properties.Settings.Default.ApplicationRegistryKey);
                if (regKey == null)
                    throw new Exception();
                this._googleCalendarName = (string)regKey.GetValue(VSTO.Properties.Settings.Default.CalendarSettings_CalendarName);
                daysBefore = (int)regKey.GetValue(VSTO.Properties.Settings.Default.CalendarSettings_RangeBefore);
                daysAfter = (int)regKey.GetValue(VSTO.Properties.Settings.Default.CalendarSettings_RangeAfter);

                // set privacy flag
                this.privacy = (int)regKey.GetValue(VSTO.Properties.Settings.Default.Privacy) == 1;
            }
            // in case no registry key exist or some registry values can't be loaded default values are used
            catch (Exception)
            { }

            // set start and end of the synchronization range
            this._rangeStart = DateTime.Now.AddDays(-daysBefore).Date;
            this._rangeEnd = DateTime.Now.AddDays(daysAfter + 1).Date;

        }

        protected override void UpdateGoogleItem(ItemMatcher pair)
        {
#if DEBUG
            Logger.Log("UpdateGoogleItem", EventType.Debug);
            var outlookItem = pair.Outlook;
            var googleItem = pair.Google;
#endif
            try
            {
                if (pair.SyncAction.Action == Action.Create)
                {
                    /// Here new Google event is created and saved into Google calendar
                    /// then it will be updated according to Outlook's original
                    /// If something wrong will happen during updating newly created Google event will be deleted
                    pair.Google = new Event()
                    {
                        ColorId = (string)Utilities.GetRegistryValue(VSTO.Properties.Settings.Default.RegistryKey_GoogleColorID),
                        End = new EventDateTime(),
                        ExtendedProperties = new Event.ExtendedPropertiesData()
                        {
                            Shared = new Dictionary<string, string>()
                        },
                        Location = "",
                        Reminders = new Event.RemindersData()
                        {
                            //Overrides = new 
                            UseDefault = false
                        },
                        Start = new EventDateTime()
                    };
                    //pair.Google.Service = this.CalendarService;
                    GoogleUtilities.SetOutlookID(pair.Google, OutlookUtilities.GetItemID(pair.Outlook));
                    try
                    {
                        Logger.Log("Setting event fields", EventType.Debug);
                        foreach (var fieldHandler in this._fieldHandlers)
                        {
                            fieldHandler.Setter(pair.Google, (Outlook.AppointmentItem)pair.Outlook, Target.Google);
                        }
                        Logger.Log("Trying to create Google event", EventType.Debug);
                        pair.Google = this.CalendarService.Events.Insert(pair.Google, this._googleCalendar.Id).Execute();
                        Logger.Log(string.Format("New Google event with eventId {0} is created", pair.Google.Id), EventType.Debug);
                        Logger.Log(Newtonsoft.Json.JsonConvert.SerializeObject(pair.Google), EventType.Debug);
                        /// When item is created it has an ID. 
                        /// If it's recurrent all recurrence instances are created. 
                        /// This gives a possibility to set recurrence exceptions.
                        /// Earlier exceptions can't be set
                        this.SetRecurrenceExceptions((Outlook.AppointmentItem)pair.Outlook, pair.Google);
                        /// As last point we set newly created Google event's ID to Outlook's item
                        OutlookUtilities.SetGoogleID(pair.Outlook, GoogleUtilities.GetItemID(pair.Google));
                        this.SaveOutlookItem(pair.Outlook);
                        this._syncResult.CreatedItems++;
                    }
                    catch (Exception exc)
                    {
                        ErrorHandler.Handle(exc);
                        if (!string.IsNullOrEmpty(pair.Google.Id)) {
                            Logger.Log(string.Format("The item '{0}' was created with ID {1} but failed to be initialized. Cleaning up...", pair.Google.Summary, pair.Google.Id), EventType.Information);
                            this._googleBatchRequest.Queue<Event>(this.CalendarService.Events.Delete(
                                this._googleCalendar.Id, pair.Google.Id), BatchCallback);
                        }
                        this._syncResult.ErrorItems++;
                    }
                }
                else if (pair.SyncAction.Action == Action.Delete)
                {
                    //pair.Google.Delete();
                    this._googleBatchRequest.Queue<Event>(this.CalendarService.Events.Delete(
                        this._googleCalendar.Id, pair.Google.Id), BatchCallback);
                    this._syncResult.DeletedItems++;
                }
                else if (pair.SyncAction.Action == Action.Update)
                {

                    /// Get list of field setters for fields, which differ
                    var fieldSetters = this._fieldHandlers.
                        Where(fieldHandler => !(bool)fieldHandler.Comparer(pair.Google, (Outlook.AppointmentItem)pair.Outlook)).
                        Select(fieldHandler => fieldHandler.Setter).
                        ToList();
                    foreach (var fieldSetter in fieldSetters)
                    {
                        fieldSetter(pair.Google, (Outlook.AppointmentItem)pair.Outlook, Target.Google);
                    }
                    if (fieldSetters.Count() > 0) {
                        pair.Google = this.CalendarService.Events.Update(pair.Google, this._googleCalendar.Id, pair.Google.Id).Execute();
                        /// When item is updated and it's became recurrent all recurrence instances are created. 
                        /// This gives a possibility to set recurrence exceptions.
                        /// Earlier exceptions can't be set
                        this.SetRecurrenceExceptions((Outlook.AppointmentItem)pair.Outlook, pair.Google);
                    }

                    this._syncResult.UpdatedItems++;
                }
            }
            catch (Exception exc)
            {
                Logger.Log(
                    String.Format("{0} {1} {2}. {3}\r\n{4}",
                        VSTO.Properties.Resources.Error_ItemSynchronizationFailure,
                        pair.SyncAction,
                        ((Outlook.AppointmentItem)pair.Outlook).Subject,
                        "", //atom,
                        ErrorHandler.BuildExceptionDescription(exc)),
                    EventType.Error);
                this._syncResult.ErrorItems++;
            }
        }

        private void BatchCallback(Event content, RequestError error, int index, HttpResponseMessage message)
        {
            if (!message.IsSuccessStatusCode)
            {
                this._batchResult.Add(error);
            }
        }

        protected override void UpdateOutlookItem(ItemMatcher pair)
        {
#if DEBUG
            var outlookItem = pair.Outlook;
            var googleItem = pair.Google;
#endif
            try
            {
                if (pair.SyncAction.Action == Action.Create)
                {
                    var outlookEvent = (Outlook.AppointmentItem)outlookApplication.CreateItem(Outlook.OlItemType.olAppointmentItem);
                    foreach (var fieldHandler in this._fieldHandlers)
                    {
                        fieldHandler.Setter(pair.Google, outlookEvent, Target.Outlook);
                    }
                    OutlookUtilities.SetGoogleID(outlookEvent, GoogleUtilities.GetItemID(pair.Google));
                    outlookEvent.Categories = (string)Utilities.GetRegistryValue(VSTO.Properties.Settings.Default.RegistryKey_OutlookCategoryName);
                    this.SaveOutlookItem(outlookEvent);
                    if (this._outlookFolderToSyncID != outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).EntryID)
                        OutlookUtilities.TryDo(() => outlookEvent.Move(outlookNamespace.GetFolderFromID(this._outlookFolderToSyncID)));
                    GoogleUtilities.SetOutlookID(pair.Google, OutlookUtilities.GetItemID(outlookEvent));
                    var res = this.CalendarService.Events.Update(pair.Google, this._googleCalendar.Id, pair.Google.Id).Execute();
                    Marshal.ReleaseComObject(outlookEvent);
                    this._syncResult.CreatedItems++;
                }
                else if (pair.SyncAction.Action == Action.Delete)
                {
                    ((Outlook.AppointmentItem)pair.Outlook).Delete();
                    Marshal.ReleaseComObject(pair.Outlook);
                    this._syncResult.DeletedItems++;
                }
                else if (pair.SyncAction.Action == Action.Update)
                {
                    /// Get list of field setters for fields, which differ
                    var fieldSetters =
                        from fieldHandler in this._fieldHandlers
                        where !fieldHandler.Comparer(pair.Google, pair.Outlook)
                        select fieldHandler.Setter;
                    foreach (var fieldSetter in fieldSetters)
                    {
                        fieldSetter(pair.Google, pair.Outlook, Target.Outlook);
                    }
                    this.SaveOutlookItem(pair.Outlook);
                    Marshal.ReleaseComObject(pair.Outlook);
                    this._syncResult.UpdatedItems++;
                }
            }
            catch (Exception exc)
            {
                Logger.Log(String.Format("{0} {1} '{2}'. {3}", VSTO.Properties.Resources.Error_ItemSynchronizationFailure, pair.SyncAction, pair.Google.Summary, ErrorHandler.BuildExceptionDescription(exc)), EventType.Error);
                this._syncResult.ErrorItems++;
            }
        }

        protected override void SaveOutlookItem(object outlookItem)
        {
            ((Outlook.AppointmentItem)outlookItem).Save();
        }

        public override SyncResult Sync()
        {
            Logger.Log("Initializing " + this.GetType().Name, EventType.Debug);
            this.Init();
            LoadGoogleItems();
            LoadOutlookItems();
            Logger.Log(String.Format("Got {0} Google items and {1} Outlook items", this.GoogleItems.Count, this.OutlookItems.Count), EventType.Information);
            this._outlookGoogleIDsCache = new Dictionary<object, string>(OutlookItems.Count);
            Logger.Log("Comparing items", EventType.Debug);
            this.CombineItems();
            /// Because error items were marked as identical we should subtract errorous items from identical ones
            this._syncResult.IdenticalItems -= this._syncResult.ErrorItems;

            Logger.Log(String.Format("There are {0} items to update", this._itemsPairs.Count), EventType.Information);
            Logger.Log("Updating items", EventType.Debug);
            /// Update items
            foreach (var pair in this._itemsPairs)
            {
                var go = (Event)pair.Google;
                var startTime = pair.Google != null ? (pair.Google.Start.Date ?? pair.Google.Start.DateTime.Value.ToString("yyyy-MM-dd")) : "";
                Logger.Log(String.Format(
                    "Running action '{0}' on item '{1}' starting at {2}. Target: {3}",
                    pair.SyncAction.Action,
                    pair.Google == null ? ((Outlook.AppointmentItem)pair.Outlook).Subject : pair.Google.Summary,
                    pair.Google == null ? ((Outlook.AppointmentItem)pair.Outlook).Start.ToString() : startTime,
                    pair.SyncAction.Target), EventType.Debug);
                try
                {
                    if (pair.SyncAction.Target == Target.Google)
                        UpdateGoogleItem(pair);
                    else
                        UpdateOutlookItem(pair);
                }
                catch (COMException exc)
                {
                    if (exc.ErrorCode == unchecked((int)0x80010105))
                    {
                        Logger.Log(exc.Message, EventType.Error);
                        this._syncResult.ErrorItems++;
                        continue;
                    }
                    else
                        throw exc;
                }
            }
            /// Run batch update on all Google items requiring update
            this._googleBatchRequest.ExecuteAsync(CancellationToken.None).Wait();
            /// Extract errors from batch result and log it
            var errors =
                from entry in this._batchResult
                where
                    entry.Code != 200 &&
                    entry.Code != 201
                select entry.Message;
            foreach (var error in errors)
                Logger.Log(error, EventType.Error);

            return this._syncResult;
        }

        internal override void Unpair()
        {
            this.Init();
            this.LoadGoogleItems();
            var batch = new BatchRequest(this.CalendarService);
            foreach (var googleItem in this.GoogleItems)
            {
                GoogleUtilities.RemoveOutlookID(googleItem);
                batch.Queue<Event>(this.CalendarService.Events.Update(googleItem, this._googleCalendar.Id, googleItem.Id), BatchCallback);
            }
            batch.ExecuteAsync(CancellationToken.None).Wait();

            this.LoadOutlookItems();
            foreach (var outlookItem in this.OutlookItems)
            {
                OutlookUtilities.RemoveGoogleID(outlookItem);
                this.SaveOutlookItem(outlookItem);
            }
        }

        protected override Event GetGoogleItemById(string id) {
            try {
                return this.CalendarService.Events.Get(this._googleCalendar.Id, id).Execute();
            } catch (GoogleApiException exc) {
                if (exc.Error.Code != 404) {
                    Logger.Log(string.Format("Couldn't get the Google item by ID '{0}'. The error is below:", id), EventType.Warning);
                    Logger.Log(exc.ToString(), EventType.Warning);
                }
                return null;
            }
        }

        protected override dynamic GetOutlookItemById(string id) {
            try {
                return this.outlookNamespace.GetItemFromID(id) as Outlook.AppointmentItem;
            } catch (COMException exc) {
                Logger.Log(string.Format("Couldn't find the Outlook item by ID '{0}'. The error is below:", id), EventType.Warning);
                Logger.Log(exc.ToString(), EventType.Warning);
                return null;
            }
        }
    }
}
