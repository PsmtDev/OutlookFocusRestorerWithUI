using System;
using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookFocusRestorerWithUI
{
    public partial class ThisAddIn
    {
        public static MailItem LastSelectedMail;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.ActiveExplorer().SelectionChange += Explorer_SelectionChange;
        }

        private void Explorer_SelectionChange()
        {
            if (Application.ActiveExplorer().Selection.Count > 0)
            {
                object item = Application.ActiveExplorer().Selection[1];
                if (item is MailItem mail)
                {
                    LastSelectedMail = mail;
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código generado por VSTO
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
