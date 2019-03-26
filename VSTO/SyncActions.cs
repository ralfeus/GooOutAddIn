using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.GoogleOutlookSync
{
    internal class SyncAction
    {
        internal Action Action { get; set; }
        internal Target Target { get; set; }

        internal SyncAction()
        { }

        internal SyncAction(Target target, Action action)
        {
            this.Action = action;
            this.Target = target;
        }

        public override string ToString()
        {
            return String.Format("{0} at {1}", this.Action, this.Target);
        }
    }

    internal enum Action
    {
        Create,
        Delete,
        Ignore,
        Update
    }

    internal enum Target
    {
        Google,
        Outlook
    }
}
