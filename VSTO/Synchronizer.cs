using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using System.Threading;
using Google.Apis.Util.Store;
//using AddinExpress.Outlook;

namespace R.GoogleOutlookSync
{
    internal class Synchronizer
	{
		public const int OutlookUserPropertyMaxLength = 32;
		public const string OutlookUserPropertyTemplate = "g/con/{0}/";
		private static object _syncRoot = new object();
        private List<ItemTypeSynchronizer> _syncBatch = new List<ItemTypeSynchronizer>(1);

        private int _skippedCount;
        public int SkippedCount
        {
            set { _skippedCount = value; }
            get { return _skippedCount; }
        }

        private int _skippedCountNotMatches;
        public int SkippedCountNotMatches
        {
            set { _skippedCountNotMatches = value; }
            get { return _skippedCountNotMatches; }
        }

        private string _propertyPrefix;
        public string OutlookPropertyPrefix
        {
            get { return _propertyPrefix; }
        }

        public string OutlookPropertyNameId
        {
            get { return _propertyPrefix + "id"; }
        }

        /*public string OutlookPropertyNameUpdated
        {
        	get { return _propertyPrefix + "up"; }
        }*/

        public string OutlookPropertyNameSynced
        {
            get { return _propertyPrefix + "up"; }
        }

        private SyncOption _syncOption = SyncOption.Merge;
        public SyncOption SyncOption
        {
            get { return _syncOption; }
            set { _syncOption = value; }
        }

        private string _syncProfile = "";
        public string SyncProfile
        {
            get { return _syncProfile; }
            set { _syncProfile = value; }
        }

        public string OutlookCalendarFolderToSyncID { get; set; }

        private bool _syncContacts;
        private UserCredential _googleCredential;
        private Outlook.Application outlookApp;
        private Outlook.NameSpace outlookNamespace { get => this.outlookApp.GetNamespace("MAPI"); }

		public void LoginToGoogle(string googleLogonAccount)
		{
			Logger.Log("Connecting to Google...", EventType.Information);
            var authorization = GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets()
                {
                    ClientId = VSTO.Properties.Settings.Default.GoogleAPI_ClientID,
                    ClientSecret = VSTO.Properties.Settings.Default.GoogleAPI_ClientSecret
                }, 
                new[] { CalendarService.Scope.Calendar }, googleLogonAccount, CancellationToken.None,
                new FileDataStore("GoogleSync"));
            this._googleCredential = authorization.Result;

            int maxUserIdLength = Synchronizer.OutlookUserPropertyMaxLength - (Synchronizer.OutlookUserPropertyTemplate.Length - 3 + 2);//-3 = to remove {0}, +2 = to add length for "id" or "up"
            string userId = this._googleCredential.UserId;
			if (userId.Length > maxUserIdLength)
				userId = userId.GetHashCode().ToString("X"); //if a user id would overflow UserProperty name, then use that user id hash code as id.
            //Remove characters not allowed for Outlook user property names: []_#
            userId.Replace("#", "").Replace("[", "").Replace("]", "").Replace("_", "");

			_propertyPrefix = string.Format(Synchronizer.OutlookUserPropertyTemplate, userId);
		}

        public void LoginToOutlook(Outlook.Application application)
        {
            this.outlookApp = application;
        }

        public void LogoffGoogle()
        {            
            //_contactsRequest = null;            
        }

        private Outlook.Items GetOutlookItems(Outlook.OlDefaultFolders outlookDefaultFolder)
        {
            Outlook.MAPIFolder mapiFolder = this.outlookNamespace.GetDefaultFolder(outlookDefaultFolder);
            try
            {
                return mapiFolder.Items;
            }
            finally
            {
                if (mapiFolder != null)
                    Marshal.ReleaseComObject(mapiFolder);
                mapiFolder = null;
            }
        }

		public void Sync()
		{
            //lock (_syncRoot)
            //{

            this.Initialize();

            var result = new CalendarSynchronizer(
                this._googleCredential,
                this.outlookApp,
                this.OutlookCalendarFolderToSyncID
                ).Sync();

            Logger.Log(String.Format("Synchronization is finished\r\n{0}", result), EventType.Information);
        }

        private string EscapeXml(string xml)
		{
			string encodedXml = System.Security.SecurityElement.Escape(xml);
			return encodedXml;
		}

        internal void Initialize()
        {
            // not sure I need it
            //this.Result = new SyncResult();
            this._syncBatch.Clear();
        }

        internal void Unpair()
        {
            this.Initialize();
            new CalendarSynchronizer(
                    this._googleCredential,
                    this.outlookApp,
                    this.OutlookCalendarFolderToSyncID
                    ).Unpair();
        }
    }
}