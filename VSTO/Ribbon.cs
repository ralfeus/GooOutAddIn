using Microsoft.Office.Core;
using R.GoogleOutlookSync;
using stdole;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace VSTO
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        private ThisAddIn host;
        private BackgroundWorker worker = new BackgroundWorker();
        private Outlook.NameSpace OutlookNamespace { get => host.Application.GetNamespace("MAPI"); }
        private Outlook.Application OutlookApplication { get => host.Application; }
        private IRibbonUI ribbon;
        private bool syncButtonPressed;

        public Ribbon(ThisAddIn host)
        {
            this.host = host;
            this.worker.WorkerSupportsCancellation = true;
            this.worker.DoWork += BackgroundJob;
            this.worker.RunWorkerCompleted += this.Worker_RunWorkerCompleted;
        }

        public void LoadRibbon(IRibbonUI ribbon) {
            this.ribbon = ribbon;
        }

        public bool GetPressed(IRibbonControl control) {
            return this.syncButtonPressed;
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
            this.syncButtonPressed = false;
            this.ribbon.Invalidate();
        }


        string IRibbonExtensibility.GetCustomUI(string RibbonID)
        {
            return Properties.Resources.Ribbon;
        }

        public void SyncGoogle(IRibbonControl control, bool pressed) {
            if (pressed) {
                Utilities.Notify("Start synchronizing", ToolTipIcon.Info);
                //this.worker.RunWorkerAsync();
                /// Needed for debugging
                Synchronize();
            } else {
                Utilities.Notify("Stop synchronizing", ToolTipIcon.Info);
                this.worker.CancelAsync();
            }
            this.syncButtonPressed = pressed;
        }

        private void BackgroundJob(object sender, DoWorkEventArgs e) {
            try {
                var intervalMinutes = (int)Utilities.GetRegistryValue(Properties.Settings.Default.RegistryKey_SyncInterval);
                Logger.Log(string.Format("Will run every {0} minutes", intervalMinutes), EventType.Information);
                while (!this.worker.CancellationPending) {
                    Synchronize();
                    Thread.Sleep(intervalMinutes * 60 * 1000);
                }
            } catch (Exception exc) {
                ErrorHandler.Handle(exc);
            }
        }

        private void Synchronize() { 
            var synchronizer = new Synchronizer();
            synchronizer.LoginToGoogle((string)Utilities.GetRegistryValue(Properties.Settings.Default.RegistryKey_GoogleAccount));
            synchronizer.LoginToOutlook(this.OutlookApplication);
            var firstTry = true;
            while (Utilities.GetRegistryValue(Properties.Settings.Default.RegistryKey_OutlookFolderID) == null && firstTry) {
                if (firstTry) {
                    ChooseFolder(null);
                    firstTry = false;
                }
            }
            if (Utilities.GetRegistryValue(Properties.Settings.Default.RegistryKey_OutlookFolderID) != null) { 
                synchronizer.OutlookCalendarFolderToSyncID = (string)Utilities.GetRegistryValue(Properties.Settings.Default.RegistryKey_OutlookFolderID);
                synchronizer.Sync();
            } else {
                throw new Exception("No Outlook folder to synchronize is chosen");
            }
        }

        public void ChooseFolder(IRibbonControl control) {
            var folder = this.OutlookNamespace.PickFolder();
            if (folder != null) {
                try {
                    Utilities.SetRegistryValue(Properties.Settings.Default.RegistryKey_OutlookFolderID, folder.EntryID);
                } catch (Exception e) {
                    ErrorHandler.Handle(e);
                }
            }
        }

        public IPictureDisp LoadImage(string imageName)
            => PictureConverter.IconToPictureDisp((Icon)Properties.Resources.ResourceManager.GetObject(imageName));

        public void Unpair(IRibbonControl control) {
            var synchronizer = new Synchronizer();
            synchronizer.LoginToGoogle((string)Utilities.GetRegistryValue(Properties.Settings.Default.RegistryKey_GoogleAccount));
            synchronizer.LoginToOutlook(this.OutlookApplication);
            var firstTry = true;
            while (Utilities.GetRegistryValue(Properties.Settings.Default.RegistryKey_OutlookFolderID) == null && firstTry) {
                if (firstTry) {
                    ChooseFolder(null);
                    firstTry = false;
                }
            }
            if (Utilities.GetRegistryValue(Properties.Settings.Default.RegistryKey_OutlookFolderID) != null) {
                synchronizer.OutlookCalendarFolderToSyncID = (string)Utilities.GetRegistryValue(Properties.Settings.Default.RegistryKey_OutlookFolderID);
                synchronizer.Unpair();
            } else {
                throw new Exception("No Outlook folder to unpair is chosen");
            }
        }
    }
}
