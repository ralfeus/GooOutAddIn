using Outlook = Microsoft.Office.Interop.Outlook;
using System.Management;
using System;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using R.GoogleOutlookSync;

namespace VSTO.Properties 
{
    
    
    // This class allows you to handle specific events on the settings class:
    //  The SettingChanging event is raised before a setting's value is changed.
    //  The PropertyChanged event is raised after a setting's value is changed.
    //  The SettingsLoaded event is raised after the setting values are loaded.
    //  The SettingsSaving event is raised before the setting values are saved.
    internal sealed partial class Settings {
        
        public Settings() {
            // // To add event handlers for saving and changing settings, uncomment the lines below:
            //
            // this.SettingChanging += this.SettingChangingEventHandler;
            //
            // this.SettingsSaving += this.SettingsSavingEventHandler;
            //
        }
        
        private void SettingChangingEventHandler(object sender, System.Configuration.SettingChangingEventArgs e) {
            // Add code to handle the SettingChangingEvent event here.
        }
        
        private void SettingsSavingEventHandler(object sender, System.ComponentModel.CancelEventArgs e) {
            // Add code to handle the SettingsSaving event here.
        }

        /// <summary>
        /// Contains method of items synchronization
        /// </summary>
        internal SyncOption SynchronizationOption { get; set; }

        //private Outlook.Application _outlookApplication = null;
        ///// <summary>
        ///// Outlook application instance. Used for creating new Outlook items
        ///// </summary>
        //internal Outlook.Application OutlookApplication
        //{
        //    get
        //    {
        //        if (this._outlookApplication == null)
        //        {
        //            Logger.Log("Connecting to Outlook...", EventType.Information);
        //            this._outlookApplication = OutlookUtilities.CreateOutlookInstance();
        //            /// Store Outlook version
        //            this.OutlookVersion = Convert.ToInt32(this._outlookApplication.Version.Substring(0, this._outlookApplication.Version.IndexOf('.')));

        //        }
        //        return this._outlookApplication;
        //    }
        //    set
        //    {
        //        if (value == null)
        //        {
        //            Logger.Log("Disconnecting from Outlook...", EventType.Debug);
        //            try
        //            {
        //                this.OutlookNamespace = null;
        //                Marshal.ReleaseComObject(this._outlookApplication);
        //                this._outlookApplication = null;
        //                Logger.Log("Disconnected from Outlook", EventType.Debug);
        //            }
        //            finally
        //            {
        //            }
        //        }
        //        else
        //            throw new NotImplementedException("Can't set directly value to property OutlookApplication");
        //    }
        //}

        //private Outlook.NameSpace _outlookNamespace;
        //internal Outlook.NameSpace OutlookNamespace
        //{
        //    get
        //    {
        //        if (!OutlookUtilities.IsAliveOutlook(this._outlookNamespace))
        //            this._outlookNamespace = this.OutlookApplication.GetNamespace("MAPI");
        //        return this._outlookNamespace;
        //    }
        //    set
        //    {
        //        if (value == null)
        //        {
        //            this.OutlookNamespace.Logoff();
        //            Marshal.ReleaseComObject(this._outlookNamespace);
        //            this._outlookNamespace = null;
        //        }
        //        else
        //            throw new NotImplementedException("Can't set directly value to property OutlookNamespace");
        //    }
        //}

        /// <summary>
        /// Outlook application major version. Necessary for features distinguishing
        /// </summary>
        internal int OutlookVersion { get; set; }

        /// <summary>
        /// Contains information whether program is registered or trial version is not expired.
        /// If application is registered property is always true
        /// </summary>
        internal bool ApplicationAllowedToRun {get;set;}// get{return true;} set{} }

        /// <summary>
        /// Contains CPU serial number
        /// </summary>
        private long _cpuSerial = 0;
        public long CPUSerial
        {
            get
            {
                try
                {
                    if (this._cpuSerial == 0)
                    {
                        var scope = new ManagementScope(@"\\.\root\cimv2");
                        scope.Connect();

                        ManagementObject wmiClass = new ManagementObject(scope, new ManagementPath("Win32_Processor.DeviceID=\"CPU0\""), new ObjectGetOptions());

                        this._cpuSerial = Convert.ToInt64((string)wmiClass.Properties["ProcessorId"].Value, 16);
                    }
                }
                catch (Exception exc)
                {
                    System.Windows.Forms.MessageBox.Show(exc.Message, exc.GetType().ToString());
                    if (exc.InnerException != null)
                        System.Windows.Forms.MessageBox.Show(exc.InnerException.Message, exc.InnerException.GetType().ToString());
                    this._cpuSerial = 0;
                }
                return this._cpuSerial;
            }
        }

        private string _outlookIDInGoogleItem = "OutlookItemID";
        public string ExtendedPropertyName_OutlookIDInGoogleItem
        {
            get { return this.CPUSerial.ToString() + "/" + this._outlookIDInGoogleItem; }
        }
    }
}
