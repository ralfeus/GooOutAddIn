using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.ComponentModel;
using System.Threading;
using Microsoft.Office.Core;
using System.Reflection;
using System.Drawing;

namespace VSTO
{
    public partial class ThisAddIn
    {
        private NotifyIcon icon;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            icon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Visible = true
            };

            //using (var worker = new BackgroundWorker()) {
            //    worker.DoWork += Worker_DoWork;
            //    worker.RunWorkerAsync();
            //}
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            while (true)
            {
                this.icon.ShowBalloonTip(3000, "AddIn", DateTime.Now.ToLongTimeString(), ToolTipIcon.Info);
                Thread.Sleep(10000);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion


        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon(this);
        }
    }
}
