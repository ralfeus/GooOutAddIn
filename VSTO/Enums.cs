using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.GoogleOutlookSync
{
    internal enum ComparisonResult
    {
        Identical,
        SameButChanged,
        Different
    }

    internal enum Field
    {
        AllDayEvent,
        Attachments,
        Attendees,
        Description,
        Location,
        Reccurent,
        Reminder,
        ShowTimeAs,
        Subject,
        Time,
        TimeZone,
        Visibility,
        Body
    }

    internal enum SyncOption
    {
        /// <summary>
        /// In case items differ freshest item will prevail
        /// </summary>
        Merge,
        /// <summary>
        /// In case items differ user will be asked which item prevails
        /// </summary>
        /// So far it's replaced with Merge option where possible
        //MergePrompt,
        /// <summary>
        /// In case items differ Google's version will be replaced with Outlook's one
        /// </summary>
        //MergeOutlookWins,
        /// <summary>
        /// In case items differ Outlook's version will be replaced with Google's one
        /// </summary>
        //MergeGoogleWins,
        /// <summary>
        /// No Outlook items are modified, no Google items are created in Outlook
        /// Outlook items are created in Google and Google items are updated with Outlook's ones
        /// </summary>
        OutlookToGoogleOnly,
        /// <summary>
        /// No Google items are modified, no Outlook items are created in Google
        /// Google items are created in Outlook and Outlook items are updated with Google's ones
        /// </summary>
        GoogleToOutlookOnly,
    }

    // Recurrency related enums
    /// <summary>
    /// Frequency of recurrent event
    /// </summary>
    internal enum RecurrenceFrequency
    {
        Daily = 0,
        Weekly = 1,
        Monthly = 2,
        MonthlyNth = 3,
        Yearly = 5
    }

    [Flags]
    internal enum DayOfWeek
    {
        Sunday = 1,
        Monday = 2,
        Tuesday = 4,
        Wednesday = 8,
        Thursday = 16,
        Friday = 32,
        Saturday = 64
    }

    internal enum EndBy
    {
        NoEnd,
        OccurencesCount,
        Date
    }

    internal enum AttendeeStatus
    {
        Tentative = 2,
        Accepted = 3,
        Declined = 4,
        NoResponse = 5
    }

    internal enum AttendeeType
    {
        Required,
        Optional,
        Organizer,
        Resource
    }
}
