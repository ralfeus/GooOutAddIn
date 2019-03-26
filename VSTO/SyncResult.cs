using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.GoogleOutlookSync
{
    class SyncResult
    {
        public int IdenticalItems { get; set; }
        public int CreatedItems { get; set; }
        public int DeletedItems { get; set; }
        public int UpdatedItems { get; set; }
        public int ErrorItems { get; set; }
        public bool Succeeded { get; set; }

        public static SyncResult operator +(SyncResult x, SyncResult y)
        {
            var res = new SyncResult();
            res.IdenticalItems = x.IdenticalItems + y.IdenticalItems;
            res.CreatedItems = x.CreatedItems + y.CreatedItems;
            res.DeletedItems = x.DeletedItems + y.DeletedItems;
            res.UpdatedItems = x.UpdatedItems + y.UpdatedItems;
            res.ErrorItems = x.ErrorItems + y.ErrorItems;
            return res;
        }

        public override string ToString()
        {
            return string.Format("Synchronization result:\r\n\tIdentical items:\t{0}\r\n\tCreated items:\t{1}\r\n\tDeleted items:\t{2}\r\n\tUpdatedItems:\t{3}\r\n\tError items:\t{4}",
                this.IdenticalItems,
                this.CreatedItems,
                this.DeletedItems,
                this.UpdatedItems,
                this.ErrorItems);
        }
    }
}
