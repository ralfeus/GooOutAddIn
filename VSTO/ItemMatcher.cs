using Google.Apis.Calendar.v3.Data;

namespace R.GoogleOutlookSync
{
    internal class ItemMatcher 
    {
        internal Event Google { get; set; }
        internal object Outlook { get; set; }
        internal SyncAction SyncAction { get; set; }

        internal ItemMatcher(Event googleItem, object outlookItem)
        {
            this.Google = googleItem;
            this.Outlook = outlookItem;
            this.SyncAction = new SyncAction();
        }
    }
}
