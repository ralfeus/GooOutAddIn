using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text.RegularExpressions;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Calendar.v3;
using Google.Apis.Requests;
using Google.Apis.Services;

namespace R.GoogleOutlookSync
{
    internal abstract class ItemTypeSynchronizer
    {
        protected Dictionary<object, string> _outlookGoogleIDsCache;

        protected IEnumerable<FieldHandlers> _fieldHandlers;
        protected IList<ItemMatcher> _itemsPairs;
        protected IList<Event> GoogleItems { get; set; }
        protected IList<object> OutlookItems { get; set; }

        protected BaseClientService _googleService;
        protected BatchRequest _googleBatchRequest;
        protected Outlook.Application outlookApplication;
        protected Outlook.NameSpace outlookNamespace { get => this.outlookApplication.GetNamespace("MAPI"); }
        protected IList<RequestError> _batchResult;
        protected SyncResult _syncResult = new SyncResult();
        /// <summary>
        /// Identifies whether private fields should be synchronized 
        /// If value is true private fields aren't synchronized
        /// Privacy of fields is defined by Private attribute of corresponding comparer and setter
        /// </summary>
        protected bool privacy = false;

        /// <summary>
        /// Represents folder in Outlook mailbox containing items for synchronization
        /// </summary>
        protected string _outlookFolderToSyncID;


        protected void Init()
        {
            this.GoogleItems = new List<Event>();
            this.OutlookItems = new List<object>();
            this._googleBatchRequest = new BatchRequest(this._googleService);
            this._batchResult = new List<RequestError>();
            //this.InitFieldHandlers();
        }

        protected void InitFieldHandlers() 
        {
            /// 1. Get methods, which current instance have. Include non-public methods too 
            ///   (handlers are private since they are not used anywhere else)
            var methods = this.GetType().GetMethods(BindingFlags.Instance | BindingFlags.NonPublic);
            /// 2. Filter only those methods, for which following condition is true:
            ///   Method has at least one attribute FieldComparer for comparer and FieldSetter for setter accordingly
            ///   In case comparer has no correspondent setter the exception will be thrown

            /// Get all comparer methods
            var comparerMethods =
                from comparerMethod in methods
                where
                    (comparerMethod.GetCustomAttributes(false).FirstOrDefault(attr => attr is FieldComparerAttribute) != null) &&
                    (!this.privacy || (comparerMethod.GetCustomAttributes(false).FirstOrDefault(attr => attr is PublicAttribute) != null))
                select comparerMethod;
            ///// Get all getter methods
            //var getterMethods =
            //    from getterMethod in methods
            //    where
            //        getterMethod.GetCustomAttributes(false).FirstOrDefault(attr => attr is FieldGetterAttribute) != null
            //    select getterMethod;
            /// Get all setter methods
            var setterMethods =
                from setterMethod in methods
                where
                    (setterMethod.GetCustomAttributes(false).FirstOrDefault(attr => attr is FieldSetterAttribute) != null) &&
                    (!this.privacy || (setterMethod.GetCustomAttributes(false).FirstOrDefault(attr => attr is PublicAttribute) != null))
                select setterMethod;
            /// Combine all field handle methods
            this._fieldHandlers = 
                from
                    comparerMethod in comparerMethods
                join setterMethod in setterMethods on
                    ((FieldHandlerAttribute)comparerMethod.GetCustomAttributes(false).First(attr => attr is FieldComparerAttribute)).Field equals
                    ((FieldHandlerAttribute)setterMethod.GetCustomAttributes(false).First(attr => attr is FieldSetterAttribute)).Field
                //join getterMethod in getterMethods on
                //    ((FieldHandler)comparerMethod.GetCustomAttributes(false).First(attr => attr is FieldComparer)).Field equals
                //    ((FieldHandler)getterMethod.GetCustomAttributes(false).First(attr => attr is FieldGetter)).Field
                select new FieldHandlers(
                    (Func<Event, object, bool>)Delegate.CreateDelegate(typeof(Func<Event, object, bool>), this, comparerMethod),
                    (Action<Event, object, Target>)Delegate.CreateDelegate(typeof(Action<Event, object, Target>), this, setterMethod));
#if DEBUG
            Func<Event, object, bool> test;
            foreach (var method in comparerMethods)
                test = (Func<Event, object, bool>)Delegate.CreateDelegate(typeof(Func<Event, object, bool>), this, method);
#endif
        }

        /// <summary>
        /// Creates a combined list of Google and Outlook items, which differ in some way
        /// For same but changed items it's one record.
        /// For each item having no pair it's one record per item
        /// </summary>
        protected virtual void CombineItems()
        {
            this._itemsPairs = new List<ItemMatcher>(this.GoogleItems.Count + this.OutlookItems.Count);
            bool noFurtherCheckNeeded;
            ComparisonResult itemsComparisonResult = 0;
            int counter = 0; int outlookItemsCount = this.OutlookItems.Count;
            foreach (var outlookItem in new List<object>(this.OutlookItems))
            {
                Logger.Log(String.Format("Checking Outlook item {0} of {1}", ++counter, outlookItemsCount), EventType.Debug);
                noFurtherCheckNeeded = false;

                foreach (var googleItem in new List<Event>(this.GoogleItems))
                {
                    try
                    {
                        itemsComparisonResult = this.Compare(googleItem, outlookItem);
                    }
                    catch (UnsynchronizableItemException exc)
                    {
                        Logger.Log(exc.Message + (exc.ItemType == Target.Google ? googleItem.Summary : OutlookUtilities.GetItemSubject(outlookItem)), EventType.Warning);
                        /// if already paired items is impossible to synchronized due to some reason
                        /// we treat them as identical
                        this._syncResult.ErrorItems++;
                        itemsComparisonResult = ComparisonResult.Identical;
                        Logger.Log(String.Format("While matching Google item '{0}' and Outlook item '{1}' the error has happened.", googleItem.Summary, OutlookUtilities.GetItemSubject(outlookItem)), EventType.Debug);
                    }
                    /// If items are same we just remove from both lists and forget about them
                    if (itemsComparisonResult == ComparisonResult.Identical)
                    {
                        Logger.Log(String.Format("Google item '{0}' and Outlook item '{1}' are identical.", googleItem.Summary, OutlookUtilities.GetItemSubject(outlookItem)), EventType.Debug);
                        this.GoogleItems.Remove(googleItem);
                        this.OutlookItems.Remove(outlookItem); //TODO: Release COM object
                        noFurtherCheckNeeded = true;
                        this._syncResult.IdenticalItems++;
                        break;
                    }
                    /// If items are same but changed we will store in pairs list for further synchronization
                    /// However if synchronization setting is one-way synchronization then all pairs, which have to
                    /// change source item will be ignored
                    else if (itemsComparisonResult == ComparisonResult.SameButChanged)
                    {
                        Logger.Log(String.Format("Google item '{0}' and Outlook item '{1}' are same but changed.", googleItem.Summary, OutlookUtilities.GetItemSubject(outlookItem)), EventType.Debug);
                        var itemMatcher = new ItemMatcher(googleItem, outlookItem);
                        if (VSTO.Properties.Settings.Default.SynchronizationOption == SyncOption.GoogleToOutlookOnly)
                            itemMatcher.SyncAction.Target = Target.Outlook;
                        else if (VSTO.Properties.Settings.Default.SynchronizationOption == SyncOption.OutlookToGoogleOnly)
                            itemMatcher.SyncAction.Target = Target.Google;
                        else
                            itemMatcher.SyncAction.Target = this.WhoLoses(googleItem, outlookItem);

                        itemMatcher.SyncAction.Action = Action.Update;
                        this._itemsPairs.Add(itemMatcher);

                        this.GoogleItems.Remove(googleItem);
                        this.OutlookItems.Remove(outlookItem);
                        noFurtherCheckNeeded = true;
                        break;
                    }
                }
                /// If no pair was found for the Outlook item and item is valid (IsItemValid() check) it will be added to the list alone for further synchronization
                /// If synchronization setting is OutlookToGoogleOnly only items with Target == Google will be added. No Outlook targeted items will be added
                if (!noFurtherCheckNeeded && this.IsItemValid(outlookItem))
                {
                    Event googleItem = null;
                    var itemMatcher = new ItemMatcher(null, outlookItem)
                    {
                        SyncAction = this.GetNonPairedItemAction(ref googleItem, outlookItem),
                        Google = googleItem
                    };

                    Logger.Log(String.Format("Outlook item '{0}' has no pair.", OutlookUtilities.GetItemSubject(outlookItem)), EventType.Debug);
                    Logger.Log(String.Format("Action {0} will be performed on {1}", itemMatcher.SyncAction.Action, itemMatcher.SyncAction.Target), EventType.Debug);

                    if (((itemMatcher.SyncAction.Target == Target.Google) && (VSTO.Properties.Settings.Default.SynchronizationOption != SyncOption.GoogleToOutlookOnly)) ||
                        ((itemMatcher.SyncAction.Target == Target.Outlook) && (VSTO.Properties.Settings.Default.SynchronizationOption != SyncOption.OutlookToGoogleOnly)))
                        this._itemsPairs.Add(itemMatcher);
                }

                //break;
            }
            /// All remaining valid (IsItemValid() check) Outlook items without Google pair will be added to the list alone for further synchronization
            /// If synchronization setting is GoogleToOutlookOnly only items with Target == Outlook will be added. No Google targeted items will be added
            Logger.Log(String.Format("{0} Outlook items remain unmatched", this.OutlookItems.Count), EventType.Debug);
            foreach (var googleItem in this.GoogleItems)
            {
                if (this.IsItemValid(googleItem))
                {
                    dynamic outlookItem = null;
                    var itemMatcher = new ItemMatcher(googleItem, null)
                    {
                        SyncAction = this.GetNonPairedItemAction(googleItem, ref outlookItem),
                        Outlook = outlookItem
                    };

                    Logger.Log(String.Format("Google item '{0}' has no pair.", googleItem.Summary), EventType.Debug);
                    Logger.Log(String.Format("Action {0} will be performed on {1}", itemMatcher.SyncAction.Action, itemMatcher.SyncAction.Target), EventType.Debug);

                    if (((itemMatcher.SyncAction.Target == Target.Google) && (VSTO.Properties.Settings.Default.SynchronizationOption != SyncOption.GoogleToOutlookOnly)) ||
                        ((itemMatcher.SyncAction.Target == Target.Outlook) && (VSTO.Properties.Settings.Default.SynchronizationOption != SyncOption.OutlookToGoogleOnly)))
                        this._itemsPairs.Add(itemMatcher);
                }
            }

            /// Clean up source Google and Outlook items lists
            this.GoogleItems.Clear();
            this.OutlookItems.Clear();
        }

        /// <summary>
        /// Compares IDs of Google and Outlook items. 
        /// Checks value of item ID with value of corresponding extended property of the peer item
        /// </summary>
        /// <param name="googleItem">Google item</param>
        /// <param name="outlookItem">Outlook item</param>
        /// <returns>true if IDs are same, false if IDs are different</returns>
        protected bool CompareIDs(Event googleItem, object outlookItem)
        {
            var outlookItemID = OutlookUtilities.GetItemID(outlookItem);
            var googleItemID = GoogleUtilities.GetItemID(googleItem);
            if (!this._outlookGoogleIDsCache.ContainsKey(outlookItem))
                this._outlookGoogleIDsCache.Add(outlookItem, OutlookUtilities.GetGoogleID(outlookItem));
            var outlookItemGoogleID = this._outlookGoogleIDsCache[outlookItem];
            var googleItemOutlookID = GoogleUtilities.GetOutlookID(googleItem);
            return (googleItemID == outlookItemGoogleID) || (outlookItemID == googleItemOutlookID);
        }

        /// <summary>
        /// Compares Google and Outlook items and check whether they are completely identical
        /// </summary>
        /// <param name="googleItem">Google item</param>
        /// <param name="outlookItem">Outlook item</param>
        /// <returns>true if items are identical, false if items differ</returns>
        protected abstract ComparisonResult Compare(Event googleItem, object outlookItem);

        /// <summary>
        /// Returns item by its ID regardless whether it's in synchronization scope or not
        /// </summary>
        /// <param name="id">ID of the item</param>
        protected abstract Event GetGoogleItemById(string id);

        /// <summary>
        /// Return an Outlook item by its ID regardless whether it's in synchronization scope or not
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        protected abstract dynamic GetOutlookItemById(string id);

        /// <summary>
        /// Defines synchronization action for Google item, which has no Outlook pair 
        /// If it's new item the Outlook item will be created
        /// If it's old item and Outlook item is already deleted the Google item will be deleted as well
        /// If Outlook's item is moved out of synchronization time rage it should be found and proper item updated 
        /// </summary>
        /// <param name="googleItem">Google item</param>
        /// <param name="outlookItem">Outlook item</param>
        /// <returns>Action and target this action should be performed on</returns>
        private SyncAction GetNonPairedItemAction(Event googleItem, ref dynamic outlookItem)
        {
            if (String.IsNullOrEmpty(GoogleUtilities.GetOutlookID(googleItem)))
                return new SyncAction(Target.Outlook, Action.Create);
            else
            {
                outlookItem = this.GetOutlookItemById(GoogleUtilities.GetOutlookID(googleItem));
                if (outlookItem != null && 
                    outlookItem.Parent.EntryID == this._outlookFolderToSyncID) {
                    var properties = ((Outlook.AppointmentItem)outlookItem).ItemProperties;
                    Logger.Log(String.Format(VSTO.Properties.Resources.Info_ItemOutsideOfRange, googleItem.Summary), EventType.Information);
                    return new SyncAction(this.WhoLoses(googleItem, outlookItem), Action.Update);
                }
                else
                {
                    return new SyncAction(Target.Google, Action.Delete);
                }
            }
        }

        /// <summary>
        /// Defines synchronization action for Outlook item, which has no Google pair 
        /// If it's new item the Google item will be created
        /// If it's old item and Google item is already deleted the Outlook item will be deleted as well
        /// If Google's item is moved out of synchronization time rage it should be found and proper item updated 
        /// </summary>
        /// <param name="googleItem">Google item</param>
        /// <param name="outlookItem">Outlook item</param>
        /// <returns>Action and target this action should be performed on</returns>
        private SyncAction GetNonPairedItemAction(ref Event googleItem, dynamic outlookItem)
        {
            if (String.IsNullOrEmpty(OutlookUtilities.GetGoogleID(outlookItem)))
                return new SyncAction(Target.Google, Action.Create);
            else
            {
                googleItem = this.GetGoogleItemById(OutlookUtilities.GetGoogleID(outlookItem));
                if (googleItem != null && googleItem.Status != "cancelled")
                {
                    Logger.Log(String.Format(VSTO.Properties.Resources.Info_ItemOutsideOfRange, googleItem.Summary), EventType.Information);
                    return new SyncAction(this.WhoLoses(googleItem, outlookItem), Action.Update);
                }
                else
                {
                    return new SyncAction(Target.Outlook, Action.Delete);
                }
            }
        }

        protected abstract bool IsItemValid(Event googleItem);
        protected abstract bool IsItemValid(object outlookItem);
        protected abstract void LoadGoogleItems();
        protected abstract void LoadOutlookItems();
        protected abstract void UpdateGoogleItem(ItemMatcher pair);
        protected abstract void UpdateOutlookItem(ItemMatcher pair);

        /// <summary>
        /// Defines what item - Google or Outlook should be updated
        /// First check synchronization settings.
        /// If synchronization settings allow both sides to be updated checks, at which side item is fresher
        /// </summary>
        /// <param name="googleItem">Google item</param>
        /// <param name="outlookItem">Outlook item</param>
        /// <returns></returns>
        private Target WhoLoses(Event googleItem, object outlookItem)
        {
            /// If merge master is defined explicitly by option we don't care on other possibilities
            /// Right now no other merge options but default one is available
            //if (Properties.Settings.Default.SynchronizationOption == SyncOption.MergeGoogleWins)
            //    return Targets.Outlook;
            //else if (Properties.Settings.Default.SynchronizationOption == SyncOption.MergeOutlookWins)
            //    return Targets.Google;
            //else if (Properties.Settings.Default.SynchronizationOption == SyncOption.MergePrompt)
            //    return this.GetUserTargetDecision();
            /// In case no weird synchronization option is set (default is Merge)
            /// Winner (and loser) will be defined by last modification time.
            /// Freshest item wins
            var googleLastModificationTime = this.GetLastModificationTime(googleItem);
            var outlookLastModificationTime = this.GetLastModificationTime(outlookItem);
            if (googleLastModificationTime < outlookLastModificationTime)
                return Target.Google;
            else if (outlookLastModificationTime < googleLastModificationTime)
                return Target.Outlook;
            else
                throw new CannotDefineSynchronizationTargetException();
        }

        protected abstract DateTime GetLastModificationTime(Event googleItem);
        protected virtual DateTime GetLastModificationTime(object outlookItem)
        {
            return OutlookUtilities.GetLastModificationTime(outlookItem);
        }

        protected Outlook.Items GetOutlookItems(string outlookFolderID)
        {
            Logger.Log("Getting Outlook folder object", EventType.Debug);
            Outlook.MAPIFolder mapiFolder = this.outlookNamespace.GetFolderFromID(outlookFolderID);
            try
            {
                Logger.Log("Getting Outlook items collection", EventType.Debug);
                return mapiFolder.Items;
            }
            finally
            {
                if (mapiFolder != null)
                    Marshal.ReleaseComObject(mapiFolder);
                mapiFolder = null;
            }
        }

        protected abstract void SaveOutlookItem(object outlookItem);
        public abstract SyncResult Sync();
        internal abstract void Unpair();
    }
}