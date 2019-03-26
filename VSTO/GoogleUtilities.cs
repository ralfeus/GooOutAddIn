using System;
using System.Linq;
using System.Net;
using Google.Apis.Calendar.v3.Data;
using System.Collections.Generic;

namespace R.GoogleOutlookSync
{
    internal static class GoogleUtilities
    {
        internal static AttendeeType GetAttendeeType(EventAttendee attendee)
        {
            if (attendee.Optional.Value)
            {
                return AttendeeType.Optional;
            } else if (attendee.Organizer.Value)
            {
                return AttendeeType.Organizer;
            } else if (attendee.Resource.Value)
            {
                return AttendeeType.Resource;
            } else
            {
                return AttendeeType.Required;
            }
        }

        internal static string GetItemID(Event item)
        {
            return item.Id;
        }

        internal static DateTime GetLastModificationTime(Event item)
        {
            return item.Updated ?? DateTime.MaxValue;
        }

        internal static string GetOutlookID(Event item)
        {
            try
            {
                var result = item.ExtendedProperties?.Shared?.First(element =>
                    element.Key == VSTO.Properties.Settings.Default.ExtendedPropertyName_OutlookIDInGoogleItem);
                return result?.Value;
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        internal static void SetOutlookID(Event item, string outlookID)
        {
            if (item.ExtendedProperties == null)
            {
                item.ExtendedProperties = new Event.ExtendedPropertiesData();
            }
            if (item.ExtendedProperties.Shared == null)
            {
                item.ExtendedProperties.Shared = new Dictionary<string, string>();
            }
            item.ExtendedProperties.Shared.Add(VSTO.Properties.Settings.Default.ExtendedPropertyName_OutlookIDInGoogleItem, outlookID);
        }

        internal static void RemoveOutlookID(Event googleItem)
        {
            var outlookIDs =
                (from outlookIDProperty in googleItem.ExtendedProperties.Shared
                where
                    outlookIDProperty.Key == VSTO.Properties.Settings.Default.ExtendedPropertyName_OutlookIDInGoogleItem
                select outlookIDProperty).ToList();
            foreach (var outlookID in outlookIDs)
                googleItem.ExtendedProperties.Shared.Remove(outlookID);
        }

        public static T TryDo<T>(Func<T> function)
        {
            System.Exception lastError = null;
            var attemptsAmount = VSTO.Properties.Settings.Default.AttemptsAmount;
            do
            {
                try
                {
                    return function();
                }
                catch (WebException exc)
                {
                    Logger.Log("During Google operation an error has occured. Error details:\r\n" + ErrorHandler.BuildExceptionDescription(exc), EventType.Debug);
                    --attemptsAmount;
                    lastError = exc;
                }
            } while (attemptsAmount > 0);
            throw new GoogleConnectionException(lastError);
        }

        public static void TryDo(System.Action function)
        {
            TryDo<object>(() => { function(); return null; });
        }
    }
}
