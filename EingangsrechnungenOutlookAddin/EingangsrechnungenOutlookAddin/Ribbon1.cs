using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.IO;

namespace EingangsrechnungenOutlookAddin
{
    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // SaveInvoice currentData = new SaveInvoice();
            //currentData.saveInvoice();
            saveInvoice();


            //MessageBox.Show(currentData.getCurrentEmailData());
        }
        private MailItem getCurrentEmailObject()
        {
            dynamic activeWindow = Globals.ThisAddIn.Application.ActiveWindow();

            try
            {
                if (activeWindow is Explorer)
                {
                    dynamic i = activeWindow.currentFolder;
                    if (activeWindow.Selection.Count > 0)
                    {
                        object selObject = activeWindow.Selection[1];
                        if (selObject is MailItem)
                        {
                            MailItem mailItem = (selObject as MailItem);
                            Marshal.ReleaseComObject(activeWindow.Selection[1]); //test reaselse
                            Marshal.ReleaseComObject(activeWindow);
                            return mailItem;
                        }
                    }
                }
                else
                {
                    Marshal.ReleaseComObject(activeWindow);
                    return activeWindow.currentitem;
                }
            }
            catch (System.Exception)
            {
                return null;
            }
            return null;
        }
        private void addLinkToEmail(string savedpath, MailItem mailItem)
        {
            mailItem.HTMLBody = savedpath + "<br>" +  mailItem.HTMLBody;
        }
        private string SaveFileTo(string initStorage, string fileName)
        {
            SaveFileDialog fd = new SaveFileDialog();
            fd.AddExtension = true;
            fd.ValidateNames = true;
            fd.FileName = fileName;
            fd.InitialDirectory = initStorage;
            fd.Filter = "PDF files|*.pdf";
            if (fd.ShowDialog() == DialogResult.OK)
                return fd.FileName;
            return "";
        }

        public void saveInvoice()
        {
                //MailItem mailObject = Globals.ThisAddIn.Application.ActiveWindow().Selection[1];
                MailItem mailObject = Globals.ThisAddIn.Application.ActiveWindow().currentitem;


            try
            {
                //MailItem mailObject = getCurrentEmailObject();
                if (mailObject != null)
                {

                    string CustomerName = "EBERTLANG";
                    //InformationFromDataBase.getCustomerNameFromDatabase(getSenderEmailAddress(mailObject));
                    if (CustomerName == "CSGClientConnection_Notfound")
                    {
                        MessageBox.Show("bitte Boxsoft Starten", "Fehler");
                        return;
                    }
                    else if (CustomerName == "EmailNotFound")
                    {
                        MessageBox.Show("Es liegt kein E-mail Adresse in der Lieferanten Ansprechpartner für die ausgewählte E-Mail-Adresse, bitte Lieferanten Ansprechpartner prüfen", "Fehler");
                        return;
                    }
                    else if (CustomerName == "ConnectionError")
                    {
                        MessageBox.Show("Keine Verbindung mit der Datenbank möglich", "Fehler");
                        return;
                    }
                    else
                    {

                        foreach (Attachment attachment in mailObject.Attachments)
                        {
                                //\\adm - stor2\TEMP\20220215_saveTemp

                            //string saveToPath = "\\\\adm-storage\\Ablage\\Alle\\_Csg\\Einkauf\\" + CustomerName + "\\Rechnungen\\" + attachment.FileName;
                            string saveToPath = "\\\\adm-stor2\\TEMP\\20220215_saveTemp\\" + attachment.FileName;

                            //if (!Directory.Exists("\\\\adm-storage\\Ablage\\Alle\\_Csg\\Einkauf\\" + CustomerName + "\\Rechnungen\\"))
                            //{
                            //      saveToPath = SaveFileTo("\\\\adm-storage\\Ablage\\Alle\\_Csg\\Einkauf\\", attachment.FileName);
                            //}

                            if (attachment.FileName.Contains(".pdf"))
                            {

                                if (string.IsNullOrEmpty(saveToPath))
                                {
                                    return;
                                }
                                attachment.SaveAsFile(saveToPath);
                                attachment.Delete();
                                addLinkToEmail(saveToPath, mailObject);
                                mailObject.SaveAs("c:\\test\\mail.msg", Outlook.OlSaveAsType.olMSG);
                            }
                        }
                        mailObject.Save();
                        Marshal.ReleaseComObject(mailObject.Attachments);
                        Marshal.ReleaseComObject(mailObject);

                    }
                }

            }
            catch (System.Exception e)
            {
                MessageBox.Show("fehler" + e.Message);
            }
        }
        private string getSenderEmailAddress(MailItem mail)
        {
            AddressEntry sender = mail.Sender;
            string SenderEmailAddress = "";

            if (sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry
                || sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
            {
                ExchangeUser exchUser = sender.GetExchangeUser();
                if (exchUser != null)
                {
                    SenderEmailAddress = exchUser.PrimarySmtpAddress;
                }
            }
            else
            {
                SenderEmailAddress = mail.SenderEmailAddress;
            }
          //  Marshal.ReleaseComObject(mail);  //realese 1
            return SenderEmailAddress;
        }
    }
}
