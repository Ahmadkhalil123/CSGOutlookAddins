using CSGSupportOutlookAddin;
using Microsoft.Office.Interop.Outlook;
using System;

namespace SupportOutlookAddIn
{
    class MailProvider
    {
        public MailProvider()
        {

        }

        public Mail GetSelectedMail()
        {

            var windowType = Globals.ThisAddIn.Application.ActiveWindow();
            if (windowType is Explorer)
            {
                Explorer explorer = windowType as Explorer;
                var currentSelection = explorer.Selection;
                if (currentSelection != null && currentSelection.Count > 0)
                {
                    if (currentSelection[1] is MailItem)
                    {
                        return new Mail((MailItem)currentSelection[1]);
                    }
                }
            }
            else if (windowType is Inspector)
            {
                Inspector inspector = windowType as Inspector;
                return new Mail(inspector.CurrentItem as MailItem);
            }
            return null;
        }
    }
}
