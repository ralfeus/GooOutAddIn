using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using Google.Apis.Calendar.v3.Data;

namespace R.GoogleOutlookSync
{
    class AttendeeComparer : IEqualityComparer<Attendee>
    {
        #region IEqualityComparer<Attendee> Members

        bool IEqualityComparer<Attendee>.Equals(Attendee x, Attendee y)
        {
            if (object.ReferenceEquals(x, y))
                return true;

            return
                x.Email == y.Email &&
                x.Required == y.Required &&
                x.Status == y.Status;
        }

        int IEqualityComparer<Attendee>.GetHashCode(Attendee obj)
        {
            return
                obj.Email.GetHashCode() ^
                obj.Required.GetHashCode() ^
                obj.Status.GetHashCode();
        }

        #endregion

        internal static bool Equals(IList<EventAttendee> googleAttendees, Outlook.Recipients outlookAttendees)
        {
            /// if Google item has recipients and their quantity is different than Outlook's one there is no reason to compare them
            if ((googleAttendees == null) || (googleAttendees.Count == 0))
            {
                return outlookAttendees.Count == 0;
            }

            foreach (Outlook.Recipient outlookAttendee in outlookAttendees)
            {
                try
                {
                    if (((Outlook.OlMeetingRecipientType)outlookAttendee.Type == Outlook.OlMeetingRecipientType.olOrganizer) ||
                        !Utilities.SMTPAddressPattern.IsMatch(outlookAttendee.Address))
                        continue;
                    var equal = false;
                    foreach (var googleAttendee in googleAttendees)
                    {
                        if (Equals(googleAttendee, outlookAttendee))
                        {
                            equal = true;
                            break;
                        }
                    }
                    if (!equal)
                        return false;
                }
                finally
                {
                    Marshal.ReleaseComObject(outlookAttendee);
                }
            }
            return true;
        }

        //internal static bool Equals(ExtensionCollection<Who> googleAttendees, vbMAPI_Recipients outlookAttendees)
        //{
        //    /// if Google item has recipients and their quantity is different than Outlook's one there is no reason to compare them
        //    if ((googleAttendees.Count > 1) && (googleAttendees.Count != outlookAttendees.Count))
        //        return false;
        //    /// if Outlook item has no recipients (we verified above Google item hasn't) return true immediately
        //    if (outlookAttendees.Count == 0)
        //        return true;

        //    foreach (vbMAPI_Recipient outlookAttendee in outlookAttendees)
        //    {
        //        try
        //        {
        //            if ((outlookAttendee.AddressEntry.Address == "Unknown") || (outlookAttendee.RecipientType == EnumRecipientType.Recipient_TO))
        //                continue;
        //            var equal = false;
        //            foreach (var googleAttendee in googleAttendees)
        //            {
        //                if (Equals(googleAttendee, outlookAttendee))
        //                {
        //                    equal = true;
        //                    break;
        //                }
        //            }
        //            if (!equal)
        //                return false;
        //        }
        //        finally
        //        {
        //            Marshal.ReleaseComObject(outlookAttendee);
        //        }
        //    }
        //    return true;
        //}

        private static bool Equals(EventAttendee googleRecipient, Outlook.Recipient outlookRecipient)
        {
            return
                googleRecipient.Email.Equals(outlookRecipient.Address, System.StringComparison.InvariantCultureIgnoreCase) &&
                (
                    (googleRecipient.Organizer.HasValue && googleRecipient.Organizer.Value && (Outlook.OlMeetingRecipientType)outlookRecipient.Type == Outlook.OlMeetingRecipientType.olOrganizer) ||
                    (googleRecipient.Optional.HasValue && googleRecipient.Optional.Value && (Outlook.OlMeetingRecipientType)outlookRecipient.Type == Outlook.OlMeetingRecipientType.olOptional) ||
                    (googleRecipient.Resource.HasValue && googleRecipient.Resource.Value && (Outlook.OlMeetingRecipientType)outlookRecipient.Type == Outlook.OlMeetingRecipientType.olResource) ||
                    ((Outlook.OlMeetingRecipientType)outlookRecipient.Type == Outlook.OlMeetingRecipientType.olRequired)
                ) && 
                (googleRecipient.ResponseStatus == ConvertTo.GoogleResponseStatus(outlookRecipient.MeetingResponseStatus));
        }

        //private static bool Equals(Who googleRecipient, vbMAPI_Recipient outlookRecipient)
        //{
        //    return
        //        (googleRecipient.Email == outlookRecipient.EmailAddress) &&
        //        (googleRecipient.Rel == ConvertTo.Google(outlookRecipient.RecipientType)) &&
        //        (googleRecipient.Attendee_Status.Value == ConvertTo.Google(outlookRecipient.).Value);
        //}
    }
}
