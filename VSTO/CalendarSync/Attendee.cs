using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Outlook;

namespace R.GoogleOutlookSync
{
    class Attendee
    {
        internal string Email { get; set; }
        internal bool Required { get; set; }
        internal OlResponseStatus Status { get; set; }

        internal Attendee(string email, int type, OlResponseStatus status)
        {
            this.Email = email;
            this.Required = (OlMeetingRecipientType)type == OlMeetingRecipientType.olRequired;
            this.Status = status;
        }

        public Attendee(string email, AttendeeType type, string responseStatus)
        {
            this.Email = email;
            this.Required = type == AttendeeType.Required;
            switch (responseStatus)
            {
                case "accepted":
                    this.Status = OlResponseStatus.olResponseAccepted;
                    break;
                case "declined":
                    this.Status = OlResponseStatus.olResponseDeclined;
                    break;
                case "tentative":
                    this.Status = OlResponseStatus.olResponseTentative;
                    break;
                default:
                    this.Status = OlResponseStatus.olResponseNotResponded;
                    break;
            }
        }
    }
}
