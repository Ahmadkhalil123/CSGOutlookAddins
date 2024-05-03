using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace Lieferanten_Dokumente_Ablegen {
    public partial class Ribbon1 {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) {
        }
        private string PDFName(MailItem mailObject) {

            string pdfname = null;
            using (Form form = new Form()) {
                Attachment pdfAttachment = mailObject.Attachments.Cast<Attachment>()
                    .FirstOrDefault(attachment => attachment.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase));
                string pdfFileName = pdfAttachment.FileName;
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(pdfFileName);
                form.Width = 350;
                form.Height = 100;
                form.Text = "PDF Name anpassen:";
                form.BackColor = SystemColors.GradientInactiveCaption;
                form.FormBorderStyle = FormBorderStyle.FixedSingle;
                form.StartPosition = FormStartPosition.CenterScreen;

                TextBox textBox = new TextBox();
                textBox.Location = new Point(10, 20);
                textBox.Width = 230;
                textBox.Text = fileNameWithoutExtension;

                Button okButton = new Button();
                okButton.Text = "OK";
                okButton.Width = 80;
                okButton.Height = textBox.Height + 1;
                okButton.DialogResult = DialogResult.OK;
                okButton.Location = new Point(textBox.Right + 1, 20);

                form.AcceptButton = okButton;
                form.Controls.Add(textBox);
                form.Controls.Add(okButton);

                DialogResult result = form.ShowDialog();
                pdfname = textBox.Text;
                if (result == DialogResult.OK) {
                    return pdfname;
                }
                return null;
            }
        }
        private string getDocumentType() {

            string selectedDocument = null;

            using (Form form = new Form()) {
                form.Width = 410;
                form.Height = 105;
                form.Text = "Dokument Auswahl:";
                form.BackColor = SystemColors.GradientInactiveCaption;
                form.FormBorderStyle = FormBorderStyle.FixedSingle;
                form.StartPosition = FormStartPosition.CenterScreen;

                Button AngebotBtn = new Button();
                AngebotBtn.Text = "Angebot";
                AngebotBtn.Size = new Size(120, 41);
                AngebotBtn.Location = new Point(12, 12);
                AngebotBtn.Click += (sender, e) =>
                {
                    form.DialogResult = DialogResult.OK; 
                    selectedDocument = "Angebote";
                    form.Close();
                };

                Button lieferscheinBtn = new Button();
                lieferscheinBtn.Text = "Lieferschein";
                lieferscheinBtn.Size = new Size(120, 41);
                lieferscheinBtn.Location = new Point(264, 12);
                lieferscheinBtn.Click += (sender, e) =>
                {
                    form.DialogResult = DialogResult.OK;
                    selectedDocument = "Lieferscheine";
                    form.Close();
                };

                Button auftragsbestätigungBtn = new Button();
                auftragsbestätigungBtn.Text = "Auftragsbestätigung";
                auftragsbestätigungBtn.Size = new Size(120, 41);
                auftragsbestätigungBtn.Location = new Point(138, 12);
                auftragsbestätigungBtn.Click += (sender, e) =>
                {
                    form.DialogResult = DialogResult.OK; 
                    selectedDocument = "Auftragsbestätigungen";
                    form.Close();
                };

                form.Controls.Add(AngebotBtn);
                form.Controls.Add(lieferscheinBtn);
                form.Controls.Add(auftragsbestätigungBtn);

                DialogResult result = form.ShowDialog();

                if (result == DialogResult.OK) {
                    return selectedDocument;
                }
                return null;
            }
        }
        private void button1_Click(object sender, RibbonControlEventArgs e) {

            MailItem mailObject;

            try {
                mailObject = Globals.ThisAddIn.Application.ActiveWindow().Selection[1];
            }
            catch (System.Exception) {

                mailObject = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;
            }

            Attachment pdfAttachment = mailObject.Attachments.Cast<Attachment>()
                  .FirstOrDefault(attachment => attachment.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase));

            if (pdfAttachment is null) {
                MessageBox.Show("Diese E-Mail enthält kein PDF");
                return;
            }

            string pdfname = PDFName(mailObject);

            if (pdfname is null) {
                return;
            }
            string documentType = getDocumentType();

            if (documentType is null) {
                return;
            }

            switch (documentType) {
                case "Angebote":
                    documentType = "AAP";
                    break;
                case "Auftragsbestätigungen":
                    documentType = "ABA";
                    break;
                case "Lieferscheine":
                    documentType = "LAP";
                    break;
            }

            string senderEmail = getSenderEmailAddress(mailObject);
            string saveToPath = "";
            string initStoragePath = "\\\\adm-storage\\Ablage\\Alle\\_Csg\\Einkauf\\";

            List<string> pathFromDB = InformationFromDataBase.getCustomerPathFromDatabase(senderEmail, documentType);

            if (pathFromDB[0] == "CSGClientConnection_Notfound") {
                MessageBox.Show("Bitte Boxsoft starten", "Fehler");
                return;
            }

            else if (pathFromDB[0] == "ConnectionError") {
                MessageBox.Show("Keine Verbindung mit der Datenbank möglich", "Fehler");
                return;
            }

            else if (pathFromDB[0] == "pathnotfound") {

                saveToPath = SaveFileTo(initStoragePath, mailObject, pdfname,senderEmail);
                if (saveToPath == "")
                    return;
            }

            else if (pathFromDB.Count > 1) {

                using (Form form = new Form()) {

                    var dropDown = new ComboBox();
                    dropDown.Items.AddRange(pathFromDB.ToArray());
                    dropDown.SelectedIndex = 1;
                    dropDown.Width = dropDown.Items[dropDown.SelectedIndex].ToString().Length * 7;
                    form.Width = dropDown.Items[dropDown.SelectedIndex].ToString().Length * 7 + 30;
                    form.Height = 100;
                    form.Text = $"Kunde hat mehrere {documentType} Felder, bitte auswählen";
                    form.StartPosition = FormStartPosition.CenterScreen;
                    dropDown.Location = new Point(10, 10);

                    Button okButton = new Button();
                    okButton.Text = "OK";
                    okButton.DialogResult = DialogResult.OK;
                    okButton.Location = new Point(form.Width - 180, 35);

                    Button cancelButton = new Button();
                    cancelButton.Text = "Cancel";
                    cancelButton.DialogResult = DialogResult.Cancel;
                    cancelButton.Location = new Point(form.Width - 100, 35);

                    form.CancelButton = cancelButton;
                    form.Controls.Add(cancelButton);
                    form.AcceptButton = okButton;
                    form.Controls.Add(okButton);
                    form.Controls.Add(dropDown);
                    form.ShowDialog();

                    if (form.DialogResult == DialogResult.OK) {
                        string selectedValue = dropDown.SelectedItem.ToString();
                        saveToPath = selectedValue;
                        form.Close();
                    }
                    if (form.DialogResult == DialogResult.Cancel) {
                        form.Close();
                        return;
                    }
                }
            }

            else {
                saveToPath = pathFromDB[0];
            }

            for (int i = mailObject.Attachments.Count; i >= 1; i--) {

                if (mailObject.Attachments[i].FileName.ToLower().Contains(".pdf")) {
                    if (string.IsNullOrEmpty(saveToPath))
                        return;
                    saveToPath = saveToPath + "\\" + pdfname.Replace(" ", "_") + ".pdf";
                    mailObject.Attachments[i].SaveAsFile(saveToPath);
                    string hyperlink = $"<a href=\"{saveToPath}\">{saveToPath}</a>";
                    mailObject.HTMLBody = $"<p>{hyperlink}</p><br>{mailObject.HTMLBody}";
                    mailObject.Attachments[i].Delete();
                }

                try {
                    mailObject.Display();
                    mailObject.Close(OlInspectorClose.olSave);
                    Clipboard.SetText(saveToPath);
                }
                catch (System.Exception msg) {
                    mailObject.Close(OlInspectorClose.olDiscard);
                    System.Threading.Thread.Sleep(1000);
                    trytosaveAgain(mailObject, saveToPath, pdfname);
                }

            }

            Marshal.ReleaseComObject(Globals.ThisAddIn.Application.ActiveWindow());
            Marshal.ReleaseComObject(mailObject);

         
        }

        private string SaveFileTo(string initStorage, MailItem mailObject, string pdfname, string senderEmail) {

            string CustomerName = InformationFromDataBase.getCustomerNameFromDatabase(senderEmail);
            if (Directory.Exists(initStorage + CustomerName + "\\")) {
                initStorage = initStorage + CustomerName + "\\";
            }

            SaveFileDialog fd = new SaveFileDialog();
            fd.AddExtension = true;
            fd.Title = "Der Pfad konnte nicht ermittelt werden, bitte wählen Sie einen Pfad aus";
            fd.ValidateNames = true;
            fd.FileName = pdfname;
            fd.InitialDirectory = initStorage;
            fd.Filter = "PDF files|*.pdf";
            if (fd.ShowDialog() == DialogResult.OK)
                return Path.GetDirectoryName(fd.FileName);
            return "";
        }
        private void trytosaveAgain(MailItem mailObject, string saveToPath, string pdfname) {

            for (int i = mailObject.Attachments.Count; i >= 1; i--) {

                int index = saveToPath.LastIndexOf("\\");
                if (index >= 0)
                    saveToPath = saveToPath.Substring(0, index) + "\\" + pdfname.Replace(" ", "_") + ".pdf"; ;

                if (mailObject.Attachments[i].FileName.Contains(".pdf")) {
                    string hyperlink = $"<a href=\"{saveToPath}\">{saveToPath}</a>";
                    mailObject.HTMLBody = $"<p>{hyperlink}</p><br>{mailObject.HTMLBody}";
                    mailObject.Attachments[i].Delete();
                }
            }

            mailObject.Display();
            mailObject.Close(OlInspectorClose.olSave);
            Clipboard.SetText(saveToPath);

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
