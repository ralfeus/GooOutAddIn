using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;

namespace R.GoogleOutlookSync
{
    internal class OutlookUtilities
    {
        /// <summary>
        /// This holds list of all types available from Outlook interop assembly.
        /// We work only with types listed there
        /// </summary>
        private static Type[] _outlookTypes;
        private static Type[] OutlookTypes
        {
            get
            {
                if (_outlookTypes == null)
                {
                    System.Reflection.Assembly outlookInterop = System.Reflection.Assembly.GetAssembly(typeof(Microsoft.Office.Interop.Outlook.MailItem));
                    _outlookTypes = outlookInterop.GetTypes();
                }
                return _outlookTypes;
            }
        }

        private static object GetItemPropertyValue(object item, string propertyName)
        {
            return TryDo<object>(() => GetItemType(item).InvokeMember(propertyName, System.Reflection.BindingFlags.GetProperty, null, item, null));
        }

        internal static string GetGoogleID(object item)
        {
            ItemProperties properties = null;
            ItemProperty property = null;
            try
            {
                properties = (ItemProperties)GetItemPropertyValue(item, "ItemProperties");
                property = properties[VSTO.Properties.Settings.Default.ExtendedPropertyName_GoogleIDInOutlookItem];
                if (property != null)
                    return property.Value.ToString();
                else
                    return null;
            }
            //catch (System.Exception)
            //{
            //    return null;
            //}
            finally
            {
                if (property != null)
                    Marshal.ReleaseComObject(property);
                if (properties != null)
                    Marshal.ReleaseComObject(properties);
            }
        }

        internal static string GetItemID(object item)
        {
            return (string)GetItemPropertyValue(item, "EntryID");
        }

        /// <summary>
        /// The type of item is unknown. Only thing is known - it's Outlook type
        /// So it's necessary first to find out what we deal with
        /// </summary>
        /// <param name="item">Outlook item</param>
        /// <returns>Type of passed item</returns>
        private static Type GetItemType(object item)
        {
            if (item is AppointmentItem)
                return typeof(AppointmentItem);
            else if (item is ContactItem)
                return typeof(ContactItem);
            else if (item is NoteItem)
                return typeof(NoteItem);
            else if (item is TaskItem)
                return typeof(TaskItem);
            else
                return null;
            
            // This way is more universal (though still requires explicit assembly specification
            // but possible performance are not clear

            //var iUnknown = Marshal.GetIUnknownForObject(item);
            //foreach (var type in OutlookTypes)
            //{
            //    var iid = type.GUID;
            //    if (!type.IsInterface || iid == Guid.Empty)
            //        continue;
            //    var iPointer = IntPtr.Zero;
            //    Marshal.QueryInterface(iUnknown, ref iid, out iPointer);
            //    if (iPointer != IntPtr.Zero)
            //    {
            //        Marshal.Release(iPointer);
            //        return type;
            //    }
            //}
            //Marshal.Release(iUnknown);
            //return null;
        }

        private static void SetItemPropertyValue(object item, string propertyName, object propertyValue)
        {
            GetItemType(item).InvokeMember(propertyName, System.Reflection.BindingFlags.SetProperty, null, item, new object[] { propertyValue });
        }

        internal static DateTime GetLastModificationTime(object item)
        {
            return (DateTime)GetItemPropertyValue(item, "LastModificationTime");
        }

        internal static void SetGoogleID(object item, string googleItemID)
        {
            ItemProperties properties = null;
            ItemProperty property = null;
            ItemProperty googleIDProperty = null;
            try
            {
                properties = (ItemProperties)GetItemPropertyValue(item, "ItemProperties");
                for (var i = 0; i < properties.Count; i++)
                {
                    property = properties[i];
                    if (property.Name == VSTO.Properties.Settings.Default.ExtendedPropertyName_GoogleIDInOutlookItem)
                    {
                        googleIDProperty = property;
                        //Marshal.ReleaseComObject(property);
                        //property = null;
                        break;
                    }
                    Marshal.ReleaseComObject(property);
                    property = null;
                }
                /// if the Outlook item is new it has no extended property for Google item ID. So we create it
                if (googleIDProperty == null)
                    googleIDProperty = properties.Add(VSTO.Properties.Settings.Default.ExtendedPropertyName_GoogleIDInOutlookItem, OlUserPropertyType.olText);
                googleIDProperty.Value = googleItemID;
            }
            catch (System.Exception exc)
            {
                ErrorHandler.Handle(exc);
            }
            finally
            {
                Marshal.ReleaseComObject(properties);
                Marshal.ReleaseComObject(googleIDProperty);
            }
        }

        internal static string GetItemSubject(object outlookItem)
        {
            if (outlookItem is AppointmentItem)
                return ((AppointmentItem)outlookItem).Subject;
            else
                throw new ArgumentException("Unknown item type");
        }

        internal static NameSpace GetOutlookNamespace(Application outlook)
        {
            NameSpace outlookNamespace = null;
            System.Exception eventualException = null;
            //Try to create new Outlook namespace 3 times, because mostly it fails the first time, if not yet running
            for (int i = 0; i < 3; i++)
            {
                try
                {
                    outlookNamespace = outlook.GetNamespace("MAPI");
                    /// Try to access the outlookNamespace to check, if it is still accessible, throws COMException, if not reachable           
                    outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                    break;  //Exit the for loop, if creating outllok application was successful
                }
                catch (COMException ex)
                {
                    eventualException = ex;
                    System.Threading.Thread.Sleep(1000 * 10);
                }
            }
            if (outlookNamespace == null)
                throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", eventualException);

            /*
            // Get default profile name from registry, as this is not always "Outlook" and would popup a dialog to choose profile
            // no matter if default profile is set or not. So try to read the default profile, fallback is still "Outlook"
            string profileName = "Outlook";
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Outlook\SocialConnector", false))
            {
                if (k != null)
                    profileName = k.GetValue("PrimaryOscProfile", "Outlook").ToString();
            }
            _outlookNamespace.Logon(profileName, null, true, false);*/

            return outlookNamespace;
        }

        internal static bool IsAliveOutlook(NameSpace _outlookNamespace)
        {
            if (_outlookNamespace == null)
                return false;
            try
            {
                _outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                return true;
            }
            catch (System.Exception)
            {
                return false;
            }
        }

        internal static void RemoveGoogleID(object outlookItem)
        {
            var properties = (ItemProperties)GetItemPropertyValue(outlookItem, "ItemProperties");
            for (int i = 0; i < properties.Count; i++)
            {
                if (properties[i].Name == VSTO.Properties.Settings.Default.ExtendedPropertyName_GoogleIDInOutlookItem)
                {
                    properties.Remove(i);
                    --i;
                }
            }
        }

        public static T TryDo<T>(Func<T> function)
        {
            var reconnections = VSTO.Properties.Settings.Default.AttemptsAmount;
            System.Exception lastError = null;
            do
            {
                var attemptsAmount = VSTO.Properties.Settings.Default.AttemptsAmount;
                do
                {
                    try
                    {
                        return function();
                    }
                    catch (COMException exc)
                    {
                        Logger.Log("During Outlook operation an error has occured. Error details:\r\n" + ErrorHandler.BuildExceptionDescription(exc), EventType.Debug);
                        lastError = exc;
                        System.Threading.Thread.Sleep(5000);
                        --attemptsAmount;
                    }
                } while (attemptsAmount > 0);
                //OutlookConnection.Disconnect();
                System.Threading.Thread.Sleep(5000);
                //OutlookConnection.Connect();
            } while (reconnections > 0);
            throw new OutlookConnectionException(lastError);
        }

        public static void TryDo(System.Action function)
        {
            TryDo<object>(() => { function(); return null; });
        }

        internal static string GetOutlookArchitecture()
        {
            for (int i = 17; i > 10; i--)
            {
                try
                {
                    var reg = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Office\\" + i.ToString() + ".0\\Outlook");
                    return reg.GetValue("Bitness").ToString();
                }
                catch { }
            }
            return "unknown";
        }
    }
}
