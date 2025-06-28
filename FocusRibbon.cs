using System;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookFocusRestorerWithUI
{
    public partial class FocusRibbon
    {
        private void FocusRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnRestoreLastMail_Click(object sender, RibbonControlEventArgs e)
        {
            if (ThisAddIn.LastSelectedMail != null)
            {
                ThisAddIn.LastSelectedMail.Display();
            }
        }
    }
}
