using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using SupportOutlookAddIn;

namespace CSGSupportOutlookAddin
{
    public partial class Ribbon1
    {
        private MailProvider mailProvider;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //Ticket
            StartCaseFormulaWithMailAddress();

        }  private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //Info
            SendCaseInformationToCustomer();
        }
      
        private void SendCaseInformationToCustomer()
        {
            var mailProvider = CreateMailProvider();
            var selectedMail = mailProvider.GetSelectedMail();

            if (selectedMail != null)
            {
                var parameter = new Dictionary<string, string>();

                parameter.Add("sender", "OutlookAddin");
                parameter.Add("action", "CaseAnswer");
                parameter.Add("emailId", selectedMail.EntryID);
                StartFormulaWithCSGDoc(parameter);
            }
        }
        private void StartCaseFormulaWithMailAddress()
        {
            var mailProvider = CreateMailProvider();
            var selectedMail = mailProvider.GetSelectedMail();

            if (selectedMail != null)
            {

                var SenderMailAdress = selectedMail.SenderEmailAddress;
                var parameter = new Dictionary<string, string>();

                parameter.Add("sender", "OutlookAddin");
                parameter.Add("action", "CreateTicket");
                parameter.Add("emailId", selectedMail.EntryID);
                parameter.Add("mailsender", SenderMailAdress);
                StartFormulaWithCSGDoc(parameter);
            }
        }

        private void StartFormulaWithCSGDoc(Dictionary<string, string> parameterList)
        {
            string CSGDocLink = "";
            var CSGDocLinkBase = @"csgdoc://cs_for_run/action=""open""&for_nr=2011&Synchron=True";

            foreach (var parameter in parameterList)
            {
                CSGDocLink = string.Format("{0}&{1}={2}", CSGDocLink, parameter.Key, parameter.Value);
            }
            CSGDocLink = CSGDocLinkBase + CSGDocLink;

            System.Diagnostics.Process.Start(CSGDocLink);
        }

        private MailProvider CreateMailProvider()
        {
            if (mailProvider == null)
            {
                mailProvider = new MailProvider();
            }
            return mailProvider;
        }

      
    }
}
