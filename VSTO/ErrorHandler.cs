using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Net.Mail;
using System.IO;
using System.Drawing;

namespace R.GoogleOutlookSync
{
    static class ErrorHandler
    {
        public static string BuildExceptionDescription(Exception exc, int deep = 0)
        {
            string identation = new String('\t', deep);
            StringBuilder result = new StringBuilder();
            result.AppendLine(identation + exc.GetType().ToString() + ": " + exc.Message + "\r\n" + identation + "\r\n" + identation + exc.StackTrace);
            if (exc.InnerException != null)
            {
                result.AppendLine(identation + "--- Inner exception ---");
                result.AppendLine(identation + BuildExceptionDescription(exc.InnerException, deep + 1));
            }
            return result.ToString();
        }

		public static void Handle(Exception exc)
        {
            string message = "Sorry, an unexpected error occured.\nProgram Version: {0}";
            message = string.Format(message, AssemblyVersion);
            try
            {
                Logger.Log(exc.Message, EventType.Error);
                Logger.Log(BuildExceptionDescription(exc), EventType.Debug);

                try { Utilities.Notify(exc.Message, ToolTipIcon.Error); }
                catch (Exception) { }
            }
            catch (Exception)
            {
                MessageBox.Show(message, Application.ProductName);
            }
        }

        public static void Handle(Exception exception, Action<bool> callback)
        {
            Handle(exception);
            //if (callback != null)
            //{
            //    if (MessageBox.Show(
            //        Properties.Resources.Confirm_SendSessionLog,
            //        Application.ProductName,
            //        MessageBoxButtons.YesNo,
            //        MessageBoxIcon.Question,
            //        MessageBoxDefaultButton.Button1) == DialogResult.Yes)
            //    {
            //        Logger.SendSessionLog(callback);
            //    }
            //}
        }

        private static string AssemblyVersion { get { return Assembly.GetExecutingAssembly().GetName().Version.ToString(); } }
    }
}
