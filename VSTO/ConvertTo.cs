using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Outlook;

namespace R.GoogleOutlookSync
{
    internal static class ConvertTo
    {
        /// <summary>
        /// Convert Outlook recipient type to Google
        /// </summary>
        /// <param name="recipientType">Outlook's recipient type</param>
        /// <returns></returns>
        internal static EventAttendee GoogleRecipientType(EventAttendee attendee, OlMeetingRecipientType recipientType)
        {
            switch (recipientType)
            {
                case OlMeetingRecipientType.olOptional:
                    attendee.Optional = true;
                    break;
                case OlMeetingRecipientType.olOrganizer:
                    attendee.Organizer = true;
                    break;
                case OlMeetingRecipientType.olResource:
                    attendee.Resource = true;
                    break;
                case OlMeetingRecipientType.olRequired:
                default:
                    attendee.Optional = attendee.Organizer = attendee.Resource = false;
                    break;
            }
            return attendee;
        }

        /// <summary>
        /// Convert Google recipient type to Outlook
        /// </summary>
        /// <param name="googleRel">Google's recipient's relation</param>
        /// <returns></returns>
        /// Probably it will be necessary to replace string with Attendee_Type type 
        internal static OlMeetingRecipientType OutlookRecipientType(EventAttendee attendee)
        {
            if (attendee.Optional.HasValue && attendee.Optional.Value) {
                return OlMeetingRecipientType.olOptional;
            } else if (attendee.Organizer.HasValue && attendee.Organizer.Value) {
                return OlMeetingRecipientType.olOrganizer;
            } else if (attendee.Resource.HasValue && attendee.Resource.Value) {
                return OlMeetingRecipientType.olResource;
            } else { 
                    return OlMeetingRecipientType.olRequired;
            }
        }

        /// <summary>
        /// Convert invitee response status from Outlook to Google
        /// </summary>
        /// <param name="outlookStatus"></param>
        /// <returns></returns>
        internal static string GoogleResponseStatus(OlResponseStatus outlookStatus)
        {
            string status;
            switch (outlookStatus)
            {
                case OlResponseStatus.olResponseAccepted:
                    status = "accepted";
                    break;
                case OlResponseStatus.olResponseDeclined:
                    status = "declined";
                    break;
                case OlResponseStatus.olResponseTentative:
                    status = "tentative";
                    break;
                default:
                    status = "needsAction";
                    break;
            }
            return status;
        }

        /// <summary>
        /// Convert Google invitee response status to Outlook's one
        /// </summary>
        /// <param name="googleStatus"></param>
        /// <returns></returns>
        internal static OlResponseStatus OutlookResponseStatus(string googleStatus)
        {
            switch (googleStatus)
            {
                case "accepted":
                    return OlResponseStatus.olResponseAccepted;
                case "declined":
                    return OlResponseStatus.olResponseDeclined;
                case "tentative":
                    return OlResponseStatus.olResponseTentative;
                default:
                    return OlResponseStatus.olResponseNotResponded;
            }
        }

        /// <summary>
        /// Convert Outlook free/busy status to Google's one.
        /// In fact Google has just two statuses: free and busy.
        /// Therefore OlBusyStatus.olFree will be converted to Google's free
        /// and rest will be converted to Google's busy
        /// </summary>
        /// <param name="outlookBusyStatus"></param>
        /// <returns></returns>
        internal static string GoogleAvailability(OlBusyStatus outlookBusyStatus)
        {
            switch (outlookBusyStatus)
            {
                case OlBusyStatus.olFree:
                    return "transparent";
                default:
                    return "opaque";
            }
        }

        /// <summary>
        /// Convert Google free/busy appearance to Outlook's one
        /// </summary>
        /// <param name="googleBusyStatus"></param>
        /// <returns></returns>
        internal static OlBusyStatus OutlookAvailability(string googleTransparency)
        {
            switch (googleTransparency)
            {
                case "transparent":
                    return OlBusyStatus.olFree;
                default: // "opaque"
                     return OlBusyStatus.olBusy;
           }
        }

        ///// <summary>
        ///// Convert any calendar event to Google's Event
        ///// </summary>
        ///// <param name="calendarEvent"></param>
        ///// <returns></returns>
        //internal static Event Google(CalendarEvent calendarEvent)
        //{
        //    return calendarEvent.ToGoogle();
        //}

        ///// <summary>
        ///// Convert any calendar event to Outlook's AppointmentItem
        ///// </summary>
        ///// <param name="calendarEvent"></param>
        ///// <returns></returns>
        //internal static AppointmentItem Outlook(CalendarEvent calendarEvent)
        //{
        //    return calendarEvent.ToOutlook();
        //}

        /// <summary>
        /// Converts Outlook item privacy/sensitivity/visibility setting to Google one
        /// </summary>
        /// <param name="olSensitivity"></param>
        /// <returns></returns>
        internal static string GoogleVisibility(OlSensitivity outlookVisibility)
        {
            switch (outlookVisibility)
            {
                case OlSensitivity.olNormal:
                    return "default";
                case OlSensitivity.olPrivate:
                    return "private";
                case OlSensitivity.olConfidential:
                    return "confidential";
                default:
                    return "public";
            }
        }

        /// <summary>
        /// Converts Google item privacy setting to Outlook's one
        /// </summary>
        /// <param name="googleVisibility"></param>
        /// <returns></returns>
        internal static OlSensitivity OutlookVisibility(string googleVisibility)
        {
            switch (googleVisibility)
            {
                case "confidential":
                    return OlSensitivity.olConfidential;
                case "private":
                    return OlSensitivity.olPrivate;
                case "public":
                default:
                    return OlSensitivity.olNormal;
            }
        }
    }
}
