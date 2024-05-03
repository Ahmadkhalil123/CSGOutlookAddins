using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using Application = Microsoft.Office.Interop.Outlook.Application;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;

namespace EingangsrechnungenOutlookAddin {
    class SaveInvoice {

        dynamic activeWindow = Globals.ThisAddIn.Application.ActiveWindow();

        private MailItem getCurrentEmailObject() {
           
            try {
                if (activeWindow is Explorer) {
                    dynamic i = activeWindow.currentFolder;
                    if (activeWindow.Selection.Count > 0) {
                        object selObject = activeWindow.Selection[1];
                        if (selObject is MailItem) {
                            MailItem mailItem = (selObject as MailItem);
                            Marshal.ReleaseComObject(activeWindow);
                            return mailItem;
                        }
                    }
                }
                else {
                    Marshal.ReleaseComObject(activeWindow);
                    return activeWindow.currentitem;
                }
            } catch (System.Exception) {
                return null;
            }
            return null;
        }
        private void addLinkToEmail(string savedpath, MailItem mailItem) {
            mailItem.HTMLBody = savedpath + "<br>" + Environment.NewLine + mailItem.HTMLBody;
        }
        private string SaveFileTo(string initStorage, string fileName) {
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

        public void saveInvoice() {

            MailItem mailObject = getCurrentEmailObject();
            if (mailObject != null) {

                string CustomerName =
                    InformationFromDataBase.getCustomerNameFromDatabase(getSenderEmailAddress(mailObject));
                if (CustomerName == "CSGClientConnection_Notfound") {
                    MessageBox.Show("bitte Boxsoft Starten", "Fehler");
                    return;
                }
                else if (CustomerName == "EmailNotFound") {
                    MessageBox.Show("Es liegt kein E-mail Adresse in der Lieferanten Ansprechpartner für die ausgewählte E-Mail-Adresse, bitte Lieferanten Ansprechpartner prüfen", "Fehler");
                    return;
                }
                else if (CustomerName == "ConnectionError") {
                    MessageBox.Show("Keine Verbindung mit der Datenbank möglich", "Fehler");
                    return;
                }
                else {

                    foreach (Attachment attachment in mailObject.Attachments) {

                        string saveToPath = "\\\\adm-storage\\Ablage\\Alle\\_Csg\\Einkauf\\" + CustomerName + "\\Rechnungen\\" + attachment.FileName;

                        if (!Directory.Exists("\\\\adm-storage\\Ablage\\Alle\\_Csg\\Einkauf\\" + CustomerName + "\\Rechnungen\\")) {

                            saveToPath = SaveFileTo("\\\\adm-storage\\Ablage\\Alle\\_Csg\\Einkauf\\", attachment.FileName);
                        }

                        if (attachment.FileName.Contains(".pdf")) {

                            if (string.IsNullOrEmpty(saveToPath)) {
                                return;
                            }
                            attachment.SaveAsFile(saveToPath);
                            attachment.Delete();
                            //addLinkToEmail(saveToPath, mailObject);
                        }
                    }
                    mailObject.Save();
                    Marshal.ReleaseComObject(mailObject);
                    activeWindow = null;
                }
            }
        }
        private string getSenderEmailAddress(MailItem mail) {
            AddressEntry sender = mail.Sender;
            string SenderEmailAddress = "";

            if (sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry
                || sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry) {
                ExchangeUser exchUser = sender.GetExchangeUser();
                if (exchUser != null) {
                    SenderEmailAddress = exchUser.PrimarySmtpAddress;
                }
            }
            else {
                SenderEmailAddress = mail.SenderEmailAddress;
            }
            return SenderEmailAddress;
        }
    }
}
