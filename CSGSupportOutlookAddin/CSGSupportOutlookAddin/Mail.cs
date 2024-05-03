using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace SupportOutlookAddIn
{
    class Mail
    {

        private MailItem _Mail;

        public Mail(MailItem mail)
        {
            _Mail = mail;

        }

        public string SenderEmailAddress
        {
            get { return _Mail.SenderEmailAddress; }
        }


        public string EntryID
        {
            get { return _Mail.EntryID; }
        }

        public string Subject
        {
            get { return _Mail.Subject; }
        }

        public string HTMLBody
        {
            get { return _Mail.HTMLBody; }
        }
    }
}
