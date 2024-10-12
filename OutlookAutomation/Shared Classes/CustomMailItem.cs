using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;
using Exception = System.Exception;
using System.Runtime.InteropServices;
using static OutlookAutomation.GlobalUtilities;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;
using File = System.IO.File;
using static OutlookAutomation.OutlookUtilities;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Newtonsoft.Json.Linq;
using System.Web.Services.Description;
using Microsoft.Office.Interop.Word;
using CheckBox = System.Windows.Forms.CheckBox;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using System.Xml.Linq;


namespace OutlookAutomation
{
    public class CustomMailItem
    {
        #region Init
        MailItem mailItem;
        ExportOptions exportOptions;
        bool nativeWordApp = false;
        public CustomMailItem(MailItem mailItem, ExportOptions exportOptions = null, Word.Application wordApp = null)
        {
            //Basic Export
            this.mailItem = mailItem;
            paths.baseFolder = exportOptions.baseFolder;

            if (exportOptions == null) { throw new ArgumentException("No export option object defined"); }
            else { this.exportOptions = exportOptions; }

            this.wordApp = wordApp;

            try { DefinePaths(); }
            catch (Exception ex) { throw new ArgumentException($"Unable to create paths.\n{ex.Message}"); }
        }

        public CustomMailItem(MailItem mailItem, ExportOptions exportOptions,
            Dictionary<string, Dictionary<string, HashSet<string>>> sortedCriteria, Word.Application wordApp = null)
        {
            //Advance Export
            this.mailItem = mailItem;

            if (exportOptions == null) { throw new ArgumentException("No export option object defined"); }
            else { this.exportOptions = exportOptions; }

            this.wordApp = wordApp;

            try { FindBaseFolders(sortedCriteria); }
            catch (Exception ex) { throw new ArgumentException($"Unable to create paths.\n{ex.Message}"); }
        }
        
        public CustomMailItem(MailItem mailItem, string filePath, ExportOptions exportOptions,
            Dictionary<string, object> hdbExportCriteria, Word.Application wordApp = null)
        {
            //HDB Export
            this.mailItem = mailItem;
            paths.baseFolder = Path.GetDirectoryName(filePath);

            if (exportOptions == null) { throw new ArgumentException("No export option object defined"); }
            else { this.exportOptions = exportOptions; }

            this.wordApp = wordApp;

            try { GenerateHDBPaths(hdbExportCriteria); }
            catch (Exception ex) { throw new ArgumentException($"Unable to create paths.\n{ex.Message}"); }
        }

        #endregion

        #region Paths
        MailPaths paths = new MailPaths();
        bool subjectFolderCreated = false;
        bool dateTimeFolderCreated = false;
        public void DefinePaths()
        {
            #region Get Subject - used for various paths
            string subject = GetSubject();
            #endregion

            #region Set or Create Subject Folder
            if (exportOptions.subjectFolder)
            {
                string subjectFolderPath = Path.Combine(paths.baseFolder, subject);
                
                paths.subjectFolder = Directory.CreateDirectory(subjectFolderPath).FullName;
                subjectFolderCreated = true;
            }
            else
            {
                paths.subjectFolder = paths.baseFolder;
            }
            #endregion

            #region Create DateTime Folder
            string sender = GetSender();

            string dateString= mailItem.SentOn.ToString("yy-MM-dd_HH-mm");
            if (exportOptions.dateTimeFolder)
            {
                // Get datetime
                //string dateString = mailItem.SentOn.ToString("yyyy-MM-dd_HH-mm");
                string thisMailFolder = Path.Combine(paths.subjectFolder, dateString + " fr " + sender);
                thisMailFolder = SanitiseFilePath(thisMailFolder);
                thisMailFolder = GetAvailableFolderName(thisMailFolder);
                string mailFolderPath = Directory.CreateDirectory(thisMailFolder).FullName;
                
                if (mailFolderPath == null)
                {
                    throw new Exception($"Unable to create folder {thisMailFolder}");
                }
                dateTimeFolderCreated = true;
                paths.thisMailFolder = mailFolderPath;
            }
            else
            {
                paths.thisMailFolder = paths.subjectFolder;
            }
            #endregion

            // Add portion that constructs the filpath and names
            #region File Names
            string fileName = $"{dateString} fr {sender}";
            paths.pdf = Path.Combine(paths.thisMailFolder, fileName + ".pdf");
            paths.rtf = Path.Combine(paths.thisMailFolder, fileName + ".rtf");
            paths.msg = Path.Combine(paths.thisMailFolder, fileName + ".msg");
            paths.html = Path.Combine(paths.thisMailFolder, fileName + ".html");
            paths.CheckPathLength();
            #endregion

        }

        private string TrimEmailSubject(string subject)
        {
            string[] prefixes = new string[] { "RE:", "FW:" ,"Fwd:" };
            foreach (var prefix in prefixes)
            {
                if (subject.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    subject = subject.Substring(prefix.Length).TrimStart(); // Trim the prefix and any whitespace
                }
            }
            subject.Trim();
            return subject;
        }

        private string GetSubject()
        {
            string subject = mailItem.Subject;
            if (subject == "" || subject == null)
            {
                subject = "No Subject";
            }
            subject = TrimEmailSubject(subject);
            subject = SanitiseFilePath(subject);
            subject = SanitiseFileName(subject);

            if (exportOptions.shortenSubject)
            {
                subject = ShortenFolderPath(subject, exportOptions.maxSubjectLength);
            }
            
            return subject;
        }

        private string GetSender()
        {
            string sender = mailItem.SenderName;
            if (sender == "" || sender == null)
            {
                sender = "No Subject";
            }
            sender = TrimEmailSubject(sender);
            sender = SanitiseFilePath(sender);
            sender = SanitiseFileName(sender);
            return sender;
        }


        #endregion

        #region Filter
        HashSet<string> baseFolders;
        private void FindBaseFolders(Dictionary<string, Dictionary<string, HashSet<string>>> sortedCriteria)
        {
            #region Filter 1 - recipients
            HashSet<string> recipients = GetRecipientSMTPAddress(mailItem);
            var validFilters = new List<Dictionary<string, HashSet<string>>>();
            
            foreach (string recipient in recipients)
            {
                string recipientLower = recipient.ToLower();
                #region Handle Full Address
                if (sortedCriteria.ContainsKey(recipientLower))
                {
                    validFilters.Add(sortedCriteria[recipientLower]);
                }
                #endregion

                #region Handle Domain only
                string domainR = recipientLower.Split('@')[1];
                domainR = "@" + domainR;
                if (sortedCriteria.ContainsKey(domainR))
                {
                    validFilters.Add(sortedCriteria[domainR]);
                }
                #endregion
            }

            if (sortedCriteria.ContainsKey(""))
            {
                validFilters.Add(sortedCriteria[""]);
            }
            Dictionary<string, HashSet<string>> validFiltersCombined = GetCombinedDictionary(validFilters);

            #endregion

            #region Filter 2 - sender
            string sender = GetSenderSMTPAddress(mailItem).ToLower();

            baseFolders = new HashSet<string>();
            if (validFiltersCombined == null) { return; }

            #region Handle Full Address
            if (validFiltersCombined.ContainsKey(sender))
            {
                foreach (string folderPath in validFiltersCombined[sender])
                {
                    if (baseFolders.Contains(folderPath)) { continue; }
                    baseFolders.Add(folderPath);
                }
            }
            #endregion

            #region Handle Domain only
            string domainS = sender.Split('@')[1];
            domainS = "@" + domainS;
            if (validFiltersCombined.ContainsKey(domainS))
            {
                foreach (string folderPath in validFiltersCombined[domainS])
                {
                    if (baseFolders.Contains(folderPath)) { continue; }
                    baseFolders.Add(folderPath);
                }
            }
            #endregion

            if (validFiltersCombined.ContainsKey(""))
            {
                foreach (string folderPath in validFiltersCombined[""])
                {
                    if (baseFolders.Contains(folderPath)) { continue; }
                    baseFolders.Add(folderPath);
                }
            }
            #endregion
        }

        private Dictionary<string, HashSet<string>> GetCombinedDictionary(List<Dictionary<string, HashSet<string>>> validFilters)
        {
            #region Simple Case no multiples
            if (validFilters.Count == 0)
            {
                return null;
            }
            else if (validFilters.Count == 1)
            {
                return validFilters[0];
            }
            #endregion

            Dictionary<string, HashSet<string>> combinedDictionary = null;
            foreach (var filterSet in validFilters)
            {
                if ( combinedDictionary == null)
                {
                    combinedDictionary = filterSet;
                    continue;
                }

                foreach (var filterType in filterSet)
                {
                    #region Doesn't already exist, add to combined
                    if (!combinedDictionary.ContainsKey(filterType.Key))
                    {
                        combinedDictionary[filterType.Key] = filterType.Value;
                        continue;
                    }
                    #endregion

                    #region Already Exist, add to hashset
                    HashSet<string> combinedHashset = combinedDictionary[filterType.Key];

                    foreach (var address in filterType.Value)
                    {
                        if (!combinedHashset.Contains(address))
                        {
                            combinedHashset.Add(address);
                        }
                    }
                    #endregion
                }
            }
            return combinedDictionary;
        }

        private HashSet<string> GetRecipientSMTPAddress(MailItem mail)
        {            
            HashSet<string> allRecipients = new HashSet<string>();
            foreach (Recipient recipient in mailItem.Recipients)
            {
                string emailAddress = GetAddressFromEntry(recipient.AddressEntry);
                if (allRecipients.Contains(emailAddress)) { continue; }
                allRecipients.Add(emailAddress);
            }
            
            return allRecipients;
        }

        private string GetAddressFromEntry(AddressEntry addressEntry)
        {
            string PR_SMTP_ADDRESS =
                @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olSmtpAddressEntry)
            {
                return addressEntry.Address;
            }

            if (!(addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry
                || addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry))
            {
                return addressEntry.PropertyAccessor.GetProperty(
                    PR_SMTP_ADDRESS) as string;
            }

            //Use the ExchangeUser object PrimarySMTPAddress
            ExchangeUser exchUser = addressEntry.GetExchangeUser();

            if (exchUser == null)
            {
                return null;
            }
            return exchUser.PrimarySmtpAddress;
        }

        private string GetSenderSMTPAddress(MailItem mail)
        {
            // Copied from https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-get-the-smtp-address-of-the-sender-of-a-mail-item

            //string PR_SMTP_ADDRESS =
            //    @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

            if (mail == null)
            {
                throw new ArgumentNullException();
            }

            if (!(mail.SenderEmailType == "EX"))
            {
                return mail.SenderEmailAddress;
            }

            AddressEntry sender = mail.Sender;
            if (sender == null)
            {
                return null;
            }
            return GetAddressFromEntry(sender);
            ////Now we have an AddressEntry representing the Sender
            //if (!(sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry
            //    || sender.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry))
            //{
            //    return sender.PropertyAccessor.GetProperty(
            //        PR_SMTP_ADDRESS) as string;
            //}

            ////Use the ExchangeUser object PrimarySMTPAddress
            //ExchangeUser exchUser = sender.GetExchangeUser();

            //if (exchUser == null)
            //{
            //    return null;
            //}
            //return exchUser.PrimarySmtpAddress;
        }
        #endregion

        #region HDB Export
        private void GenerateHDBPaths(Dictionary<string, object> hdbExportCriteria)
        {
            (string subjectFolderName, string toFromString) = FindTargetFoldersFromCriteria(hdbExportCriteria);
            DefinePathsForHDB(subjectFolderName, toFromString);
        }
        private (string subjectFolderName, string toFromString) FindTargetFoldersFromCriteria(Dictionary<string, object> hdbExportCriteria)
        {
            #region Subject determines base path
            string subject = mailItem.Subject;
            if (subject.Length == 0) { throw new ArgumentException("Subject is empty"); }

            Dictionary<string, string> subjectStrings = (Dictionary<string,string>)hdbExportCriteria["subjectStrings"];
            string subjectFolderName = "";
            foreach (KeyValuePair<string,string> entry in subjectStrings)
            {
                int first = subject.IndexOf(entry.Key);
                if (first == -1) { continue; } // Not found


                bool isInteger = Int32.TryParse(entry.Value, out int subjectLength);
                if (!isInteger) { throw new ArgumentException("Unable to parse {entry.Value} into integer");}

                if (subjectLength == 0)
                {
                    subjectFolderName = subject.Substring(first);
                    int spacePosition = subjectFolderName.IndexOf(' ');
                    subjectFolderName = subjectFolderName.Substring(0, spacePosition);
                }
                else
                {
                    subjectFolderName = subject.Substring(first, subjectLength);
                }
                
                break;
            }

            subjectFolderName = SanitiseFilePath(subjectFolderName);
            subjectFolderName = SanitiseFileName(subjectFolderName);
            #endregion

            #region Sender
            string senderName = mailItem.SenderName.ToLower();
            string senderEmail = GetSenderSMTPAddress(mailItem);
            if (senderEmail != null) { senderEmail = senderEmail.ToLower(); }
            HashSet<string> internalSenders = (HashSet<string>)hdbExportCriteria["internalSenders"];
            bool isInternalSender = false;
            if (internalSenders.Contains(senderName)) { isInternalSender = true; }
            if (senderEmail != null && internalSenders.Contains(senderEmail)) { isInternalSender = true; }
            #endregion

            #region Recipient
            string toFromString;
            string referenceName = "";
            Dictionary<string, string> externalReferenceNames = (Dictionary<string, string>)hdbExportCriteria["externalReferenceNames"];
            if (isInternalSender) // find to
            {
                (string[] toRecipientsNames, string[] toRecipientsEmailAddress) = GetOnlyToRecipientInfo(mailItem);

                //foreach(string recipientEmail in toRecipientsEmailAddress)
                for (int i = 0; i < toRecipientsEmailAddress.Length; i++)
                {
                    string recipientEmail = toRecipientsEmailAddress[i];
                    string recipientName = toRecipientsNames[i];
                    string tryReplacement = emailToBeReplaced(recipientEmail, externalReferenceNames);
                    if (tryReplacement != null)
                    {
                        referenceName = tryReplacement;
                        break;
                    }
                    else if (externalReferenceNames.ContainsKey(recipientName))
                    {
                        referenceName = externalReferenceNames[recipientName];
                        break;
                    }
                }

                if (referenceName == "") 
                {
                    LogMailItemWarning($"No matching recipient replacement found, default to first recipient.");
                    referenceName = mailItem.Recipients[1].Name;
                }
                toFromString = $" to {referenceName}";
            }
            #endregion
            
            #region External Sender
            else // find fr
            {
                string tryReplacement = emailToBeReplaced(senderEmail, externalReferenceNames);

                if (tryReplacement != null)
                {
                    referenceName = tryReplacement;
                }
                else if (externalReferenceNames.ContainsKey(senderName))
                {
                    referenceName = externalReferenceNames[senderName];
                }
                else
                {
                    referenceName = senderName;
                }
                toFromString = $" fr {referenceName}";
            }

            toFromString = SanitiseFilePath(toFromString);
            toFromString = SanitiseFileName(toFromString);
            #endregion

            return (subjectFolderName, toFromString);
        }
        private string emailToBeReplaced(string email, Dictionary<string,string> externalReferenceNames)
        {
            if (email == null) { return null; }
            string domain = email.Split('@')[1];
            domain = "@" + domain;

            if (externalReferenceNames.ContainsKey(email))
            {
                return externalReferenceNames[email];
            }
            else if (externalReferenceNames.ContainsKey(domain))
            {
                return externalReferenceNames[domain];
            }
            return null;
        }
        private string externalNameToBeReplaced(string name, Dictionary<string, string> externalReferenceNames)
        {
            if (name == null) { return null; }
            
            if (externalReferenceNames.ContainsKey(name))
            {
                return externalReferenceNames[name];
            }

            return null;
        }
        public void DefinePathsForHDB(string subjectFolderName, string toFromString)
        {
            #region Get Subject - used for various paths
            string subject = GetSubject();
            #endregion

            #region Set or Create Subject Folder
            if (exportOptions.subjectFolder)
            {
                string subjectFolderPath;
                if (subjectFolderName != "")
                {
                    subjectFolderPath = Path.Combine(paths.baseFolder, subjectFolderName);
                }
                else
                {
                    subjectFolderPath = Path.Combine(paths.baseFolder, GetSubject()); 
                }
                
                paths.subjectFolder = Directory.CreateDirectory(subjectFolderPath).FullName;
                subjectFolderCreated = true;
            }
            else
            {
                paths.subjectFolder = paths.baseFolder;
            }
            #endregion

            #region Create DateTime Folder
            string dateString = mailItem.SentOn.ToString("yy-MM-dd_HH-mm");
            string thisMailFolder = Path.Combine(paths.subjectFolder, dateString + toFromString);
            thisMailFolder = SanitiseFilePath(thisMailFolder);
            thisMailFolder = GetAvailableFolderName(thisMailFolder);
            string mailFolderPath = Directory.CreateDirectory(thisMailFolder).FullName;

            if (mailFolderPath == null)
            {
                throw new Exception($"Unable to create folder {thisMailFolder}");
            }
            dateTimeFolderCreated = true;
            paths.thisMailFolder = mailFolderPath;
            #endregion

            // Add portion that constructs the filpath and names
            #region File Names
            string fileName = $"{dateString} Email" + toFromString;
            paths.pdf = Path.Combine(paths.thisMailFolder, fileName + ".pdf");
            paths.rtf = Path.Combine(paths.thisMailFolder, fileName + ".rtf");
            paths.msg = Path.Combine(paths.thisMailFolder, fileName + ".msg");
            paths.html = Path.Combine(paths.thisMailFolder, fileName + ".html");
            paths.CheckPathLength();
            #endregion
        }

        private (string[]names ,string[]emails) GetOnlyToRecipientInfo(MailItem mailItem)
        {
            List<string> toRecipientEmails = new List<string>();
            List<string> toRecipientNames = new List<string>();
            foreach (Recipient recipient in this.mailItem.Recipients)
            {
                if (recipient.Type != (int)OlMailRecipientType.olTo) // Skip those not in "To" field
                {
                    continue;
                }
                string emailAddress = GetAddressFromEntry(recipient.AddressEntry);
                
                if (toRecipientEmails.Contains(emailAddress)) { continue; }
                toRecipientEmails.Add(emailAddress);
                toRecipientNames.Add(recipient.Name.ToLower());
            }
            return (toRecipientNames.ToArray(), toRecipientEmails.ToArray());
        }

        #endregion

        #region Export Functions
        bool exported = false;
        public void ExportWithFilters()
        {
            if (baseFolders.Count == 0) 
            {
                Beaver.LogError($"No matching filters found for email.\n" +
                    $"    Subject: {mailItem.Subject}\n" +
                    $"    Date: {mailItem.SentOn.ToString("dddd, dd MMMM yyyy h:mm tt")}\n");
                return;
            }

            foreach (string baseFolder in baseFolders)
            {
                paths.baseFolder = baseFolder;
                DefinePaths();
                Export(true);

                Beaver.LogProgress($"Email successfully exported.\n" +
                    $"    Subject: {mailItem.Subject}\n" +
                    $"    Date: {mailItem.SentOn.ToString("dddd, dd MMMM yyyy h:mm tt")}\n" +
                    $"    Folder: {baseFolder}\n");
            }
            exported = true;
        }
        public void Export(bool exportMultiple = false)
        {
            if (exportOptions.pdf || exportOptions.word ) { ExportPdfAndWord(); }
            if (exportOptions.msg) { ExportMsg(); }
            if (exportOptions.html) { ExportHtml(); }
            if (exportOptions.attachments) { ExportAttachments(); }
            if (!exportMultiple) { exported = true; } 
        }

        public void ExportPdfAndWord()
        {
            try
            {
                ExportRTF();
                if (exportOptions.pdf) { ExportWordAsPdf(); }
            }
            finally
            {
                ReleaseItems();
            }
        }

        public void ExportMsg()
        {
            mailItem.SaveAs(paths.msg, OlSaveAsType.olMSG);
        }

        public void ExportHtml()
        {
            mailItem.SaveAs(paths.html, OlSaveAsType.olHTML);
        }

        public void ExportRTF()
        {
            string path = paths.msg.Substring(0,paths.msg.Length-4);
            path += ".rtf";
            mailItem.SaveAs(path, OlSaveAsType.olRTF);
        }

        public void ExportAttachments()
        {
            if (mailItem.Attachments.Count == 0) { return; }
            string attachmentFolder = Path.Combine(paths.thisMailFolder, "Attachments");
            
            foreach (Attachment attachment in mailItem.Attachments)
            {
                if (!isMainAttachment(attachment)) { continue; }

                string fileName = "";
                try
                {
                    fileName = attachment.FileName;
                }
                catch 
                {
                    string msg = $"Unable to get filename for attachment. Please check output.\n" +
                        $"    Mail Subject: {mailItem.Subject}\n" +
                        $"    Date: {mailItem.SentOn.ToString("dddd, dd MMMM yyyy h:mm tt")}\n" +
                        $"    Attachment name: {attachment.DisplayName}\n";
                    Beaver.LogError(msg);
                    continue;
                }

                if (!Directory.Exists(attachmentFolder))
                {
                    Directory.CreateDirectory(attachmentFolder);
                }

                string attachmentPath = Path.Combine(attachmentFolder, fileName);
                attachmentPath = GetAvailableFileName(attachmentPath);
                if (attachmentPath.Length > 248) { throw new ArgumentException($"Attachment File path exceeds maximum allowable characters: {attachmentPath}."); }
                attachment.SaveAsFile(attachmentPath);
            }
        }
        #endregion

        #region Word Document

        #region Headers
        //List<(string type, string value)> headers;
        //private void GetHeaders()
        //{
        //    headers = new List<(string type, string value)>();
        //    headers.Add(("From", $"{mailItem.SenderName}<{GetSenderSMTPAddress(mailItem)}>"));
        //    headers.Add(("Sent", mailItem.SentOn.ToString("dddd, dd MMMM yyyy h:mm tt")));
        //    headers.Add(("To", mailItem.To));
        //    headers.Add(("Subject", mailItem.Subject));
        //    headers.Add(("Attachments", GetAttachmentNames()));
        //}


        #endregion
        
        Word.Document tempDoc;
        Word.Application wordApp;

        
        private void ExportWordAsPdf()
        {
            if (wordApp == null)
            {
                wordApp = new Word.Application();
                nativeWordApp = true;
            }

            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone; // Suppress alerts
            tempDoc = wordApp.Documents.Open(paths.rtf);
            
            //RetryIfBusy(() => { tempDoc = wordApp.Documents.Open(paths.rtf); });

            RetryIfBusy(() =>
            {
                tempDoc.PageSetup.TopMargin = wordApp.CentimetersToPoints(1.27f);
                tempDoc.PageSetup.BottomMargin = wordApp.CentimetersToPoints(1.27f);
                tempDoc.PageSetup.LeftMargin = wordApp.CentimetersToPoints(1.27f);
                tempDoc.PageSetup.RightMargin = wordApp.CentimetersToPoints(1.27f);
            });
            InsertHeader();
            tempDoc.ExportAsFixedFormat(paths.pdf, Word.WdExportFormat.wdExportFormatPDF);
        }

        private void InsertHeader()
        {
            Word.Range range = null;
            Word.Range tbRange = null;
            Word.Table tb = null;
            Word.Table table = null;
            try
            {
                #region Set range to edit
                range = tempDoc.Range(0, 0);
                if (tempDoc.Tables.Count > 0 && tempDoc.Tables[1].Range.Start == 0)
                {
                    tb = tempDoc.Tables[1];
                    tbRange = tb.Range;
                    tbRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    tbRange.InsertParagraphAfter();
                    tb.Range.Relocate(1);
                }
                else
                {
                    range.InsertParagraphBefore();
                }
                range.SetRange(0, 0);
                #endregion

                #region Insert Name and Border
                range.Text = Globals.ThisAddIn.Application.Session.CurrentUser.Name;
                range.InsertParagraphAfter();
                range.set_Style(Word.WdBuiltinStyle.wdStyleNormal);
                range.Bold = 1;
                range.Font.Name = "Calibri";
                range.Font.Size = 11;
                range.Font.Color = Word.WdColor.wdColorBlack;
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;

                #endregion

                //if (exportOptions.word)
                //{
                //    tempDoc.Save();
                //}
            }
            finally
            {
                if (range != null)
                {
                    Marshal.ReleaseComObject(range);
                    range = null;
                }

                if (tbRange != null)
                {
                    Marshal.ReleaseComObject(tbRange);
                    tbRange = null;
                }

                if (tb != null)
                {
                    Marshal.ReleaseComObject(tb);
                    tb = null;
                }

                if (table != null)
                {
                    Marshal.ReleaseComObject(table);
                    table = null;
                }
            }
        }

        #endregion

        #region Attachments
        //private string GetAttachmentNames()
        //{
        //    //string attachmentNames = "";
        //    //Attachments attachments = mailItem.Attachments;
        //    //foreach (Attachment attachment in attachments)
        //    //{
        //    //    if (!isMainAttachment(attachment)) { continue; }

        //    //    string fileName = "";
        //    //    try
        //    //    {
        //    //        fileName = attachment.FileName;
        //    //    }
        //    //    catch //(Exception ex)
        //    //    {
        //    //        // Skip embeded items that we can't get filenames for 
        //    //        continue;
        //    //        //fileName = attachment.DisplayName;
        //    //    }
        //    //    attachmentNames += fileName + "\n";
        //    //}

        //    //if (attachmentNames == "")
        //    //{
        //    //    attachmentNames = "None";
        //    //}
        //    //else
        //    //{
        //    //    attachmentNames = attachmentNames.Substring(0, attachmentNames.Length - 2);
        //    //}

        //    //return attachmentNames;
        //}

        private bool isMainAttachment(Attachment attachment)
        {
            if (!exportOptions.skipEmbedded) { return true; } //Skip embeded is not selected 
            var flags = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003");
            // Might have to update the proptag 0x37140003 if microsoft wants to be screwy
            if (flags != 0) { return false; }
            else { return true; }
        }
        #endregion

        #region Release Items
        public void ReleaseItems()
        {
            if (tempDoc != null)
            {
                tempDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                
                if (!exportOptions.word)
                {
                    try { File.Delete(paths.rtf); }
                    catch { }
                }
                
                Marshal.FinalReleaseComObject(tempDoc);
                tempDoc = null;
            }
            
            if (nativeWordApp && wordApp != null)
            {
                wordApp.Quit(false);
                Marshal.ReleaseComObject(wordApp);
                wordApp = null;
            }
        }
        public void DeleteEmptyFolders()
        {
            if (dateTimeFolderCreated)
            {
                DeleteFolderIfEmpty(paths.thisMailFolder);
            }

            if (subjectFolderCreated)
            {
                DeleteFolderIfEmpty(paths.subjectFolder);
            }
        }

        #endregion

        #region Move To Folder
        #region Move MailItem

        public void MoveToFolder()
        {
            GetFolders();
            if (exported)
            {
                mailItem.Move(exportFolder);
            }
            else
            {
                mailItem.Move(failedFolder);
            }
        }

        Folder exportFolder;
        Folder failedFolder;
        private void GetFolders()
        {
            Application application = Globals.ThisAddIn.Application;
            Folder rootFolder = application.ActiveExplorer().CurrentFolder as Folder;
            //Folder root = application.Session.DefaultStore.GetRootFolder() as Folder;
            //Folder baseExport = GetFolderFromFolder(rootFolder, "Exports", true);
            exportFolder = GetFolderFromFolder(rootFolder, "Exported", true);
            failedFolder = GetFolderFromFolder(rootFolder, "Failed", true);
        }


        #endregion
        #endregion

        #region Beaver
        private void LogMailItemWarning(string errorMsg)
        {
            string msg = $"Warning: Problem encountered during export.\n" +
                $"    Subject: {mailItem.Subject}\n" +
                $"    Date: {mailItem.SentOn.ToString("dddd, dd MMMM yyyy h:mm tt")}\n" +
                $"    Error Message: {errorMsg}\n";
            Beaver.LogError(msg);
        }
        #endregion
    }

    #region Mail Paths 
    class MailPaths
    {
        public MailPaths() { }
        public string baseFolder;
        public string subjectFolder;
        public string thisMailFolder;

        public string msg;
        public string pdf;
        public string html;
        public string rtf;
        public void CheckPathLength()
        {
            if (thisMailFolder.Length > 256)
            {
                throw new ArgumentException($"Folder path exceeds maximum character limit of 256\n {thisMailFolder}");
            }

            if (pdf.Length + 1 > 248)
            {
                throw new ArgumentException($"File path exceeds maximum character limit\n {msg}");
            }
        }
    }
    
    #region Archive
    //public class ExportOptions0
    //{
    //    public bool attachments;
    //    public bool html;
    //    public bool msg;
    //    public bool pdf;
    //    public bool word;


    //    public bool subjectFolder;
    //    public bool dateTimeFolder;

    //    public bool skipEmbeded;
    //    public bool shortenSubject;
    //    public int maxSubjectLength;
    //    public bool breakOnError;

    //    public string baseFolder;

    //    public bool moveItems;

    //    #region Init

    //    public ExportOptions0()
    //    {
    //        LoadFromSettings();
    //    }
    //    #region Archive
    //    public ExportOptions0(bool? exportAll = null)
    //    {
    //        if (exportAll == null) { return; }
    //        else if ((bool)exportAll)
    //        {
    //            attachments = true;
    //            html = true;
    //            msg = true;
    //            pdf = true;
    //            word = true;

    //            subjectFolder = true;
    //            dateTimeFolder = true;

    //            skipEmbeded = true;
    //            shortenSubject = true;
    //            breakOnError = true;
    //        }
    //        else
    //        {
    //            attachments = false;
    //            html = false;
    //            msg = false;
    //            pdf = false;
    //            word = false;

    //            subjectFolder = false;
    //            dateTimeFolder = false;

    //            skipEmbeded = false;
    //            shortenSubject = false;
    //            breakOnError = true;
    //        }
    //    }
    //    #endregion

    //    public void RefreshFromUserControl(
    //        bool attachments,
    //        bool html,
    //        bool msg,
    //        bool pdf,
    //        bool word,

    //        bool subjectFolder,
    //        bool dateTimeFolder,

    //        bool skipEmbeded,
    //        bool shortenSubject,
    //        string maxSubjectLength,
    //        bool breakOnError,

    //        string baseFolder,

    //        bool moveItems
    //        )
    //    {
    //        this.attachments = attachments;
    //        this.html = html;
    //        this.msg = msg;
    //        this.pdf = pdf;
    //        this.word = word;

    //        this.dateTimeFolder = dateTimeFolder;
    //        this.subjectFolder = subjectFolder;

    //        this.skipEmbeded = skipEmbeded;
    //        this.shortenSubject = shortenSubject;
    //        this.breakOnError = breakOnError;

    //        this.moveItems = moveItems;


    //        bool canParse = Int32.TryParse(maxSubjectLength, out this.maxSubjectLength);
    //        if (!canParse)
    //        {
    //            throw new ArgumentException($"Unable to parse {maxSubjectLength} into integer for maximum subject length");
    //        }

    //        this.baseFolder = baseFolder;
    //        if (!Directory.Exists(baseFolder))
    //        {
    //            throw new ArgumentException($"Warning: Base Folder path does not exist.\n{baseFolder}");
    //        }
    //        SaveToSettings();
    //    }
    //    #endregion

    //    #region Save to and Load from Settings
    //    public void SaveToSettings()
    //    {
    //        Properties.Settings settings = Properties.Settings.Default;
    //        // Should turn this into dictionary and loop
    //        settings.exportAttachments = attachments;
    //        settings.exportHtml = html;
    //        settings.exportMsg = msg;
    //        settings.exportPdf = pdf;
    //        settings.exportWord = word;

    //        settings.dateFolder = dateTimeFolder;
    //        settings.subjectFolder = subjectFolder;

    //        settings.skipEmbedded = skipEmbeded;
    //        settings.shortenSubject = shortenSubject;
    //        settings.maxSubjectLength = maxSubjectLength;
    //        settings.breakOnError = breakOnError;

    //        settings.baseFolder = baseFolder;

    //        settings.moveItems = moveItems;
    //        settings.Save();
    //    }

    //    public void LoadFromSettings()
    //    {
    //        Properties.Settings settings = Properties.Settings.Default;

    //        attachments = settings.exportAttachments;
    //        html = settings.exportHtml;
    //        msg = settings.exportMsg;
    //        pdf = settings.exportPdf;
    //        word = settings.exportWord;

    //        dateTimeFolder = settings.dateFolder;
    //        subjectFolder = settings.subjectFolder;

    //        skipEmbeded = settings.skipEmbedded;
    //        shortenSubject = settings.shortenSubject;
    //        maxSubjectLength = settings.maxSubjectLength;
    //        breakOnError = settings.breakOnError;

    //        baseFolder = settings.baseFolder;

    //        moveItems = settings.moveItems;
    //    }
    //    #endregion
    //}
    #endregion
    
    #region Export Options
    public class ExportOptions
    {
        #region Init
        Dictionary<string, CustomSetting> settingTracker = new Dictionary<string, CustomSetting>();
        public ExportOptions() { }
        public void Add(CustomSetting customSetting)
        {
            settingTracker.Add(customSetting.name, customSetting);
        }
        #endregion

        #region Save and Load
        public void SaveSettings()
        {
            Properties.Settings.Default.Save();
        }

        public void LoadSettings()
        {
            foreach (CustomSetting customSetting in settingTracker.Values) { customSetting.LoadValue(); }
            foreach (CustomSetting customSetting in settingTracker.Values) { customSetting.SubscribeToEvents(); }

        }
        #endregion

        #region Get Value
        public bool attachments
        {
            get { return ((CheckSetting)settingTracker["exportAttachments"]).Value; }
        }

        public bool html
        {
            get { return ((CheckSetting)settingTracker["exportHtml"]).Value; }
        }

        public bool msg
        {
            get { return ((CheckSetting)settingTracker["exportMsg"]).Value; }
        }

        public bool pdf
        {
            get { return ((CheckSetting)settingTracker["exportPdf"]).Value; }
        }

        public bool word
        {
            get { return ((CheckSetting)settingTracker["exportWord"]).Value; }
        }

        public bool dateTimeFolder
        {
            get { return ((CheckSetting)settingTracker["dateFolder"]).Value; }
        }

        public bool subjectFolder
        {
            get { return ((CheckSetting)settingTracker["subjectFolder"]).Value; }
        }

        public bool skipEmbedded
        {
            get { return ((CheckSetting)settingTracker["skipEmbedded"]).Value; }
        }

        public bool shortenSubject
        {
            get { return ((CheckSetting)settingTracker["shortenSubject"]).Value; }
        }

        public int maxSubjectLength
        {
            get { return ((IntBoxSetting)settingTracker["maxSubjectLength"]).Value; }
        }

        public bool breakOnError
        {
            get { return ((CheckSetting)settingTracker["breakOnError"]).Value; }
        }

        public string baseFolder
        {
            get { return ((TextBoxSetting)settingTracker["baseFolder"]).Value; }
        }

        public bool moveItems
        {
            get { return ((CheckSetting)settingTracker["moveItems"]).Value; }
        }
        #endregion
    
    }

    #region Custom Setting Class
    public abstract class CustomSetting
    {
        public string name;

        public CustomSetting(string name)
        {
            this.name = name;
        }
        public abstract void SubscribeToEvents();
        public abstract void LoadValue();
        public virtual object Value
        {
            get { return null; }
        }

    }
    public class CheckSetting : CustomSetting
    {
        public CheckBox checkBox;
        public CheckSetting(string name, CheckBox checkBox) : base(name)
        {
            this.checkBox = checkBox;
        }

        public override void SubscribeToEvents()
        {
            checkBox.CheckedChanged += new EventHandler(CheckBox_CheckedChanged);
        }

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default[name] = checkBox.Checked;
        }
        public override void LoadValue() { checkBox.Checked = (bool)Properties.Settings.Default[name]; }
        public new bool Value
        {
            get { return (bool) Properties.Settings.Default[name]; }
        }
    }

    internal class TextBoxSetting : CustomSetting
    {
        #region Init
        public TextBox textBox;
        internal TextBoxSetting(string name, TextBox textBox) : base(name)
        {
            this.textBox = textBox;
        }

        public override void SubscribeToEvents()
        {
            textBox.LostFocus += new EventHandler(textBox_LostFocus);
            textBox.KeyDown += new KeyEventHandler(textBox_KeyDown);
        }

        private void textBox_LostFocus(object sender, EventArgs e)
        {
            Properties.Settings.Default[name] = textBox.Text;
        }

        protected void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox.Parent.Focus();
            }
        }
        #endregion
        public override void LoadValue() { textBox.Text = (string)Properties.Settings.Default[name]; }
        public new string Value
        {
            get { return (string)Properties.Settings.Default[name]; }
        }
    }

    internal class IntBoxSetting : CustomSetting
    {
        #region Init
        public TextBox textBox;
        internal IntBoxSetting(string name, TextBox textBox) : base(name)
        {
            this.textBox = textBox;
        }

        public override void SubscribeToEvents()
        {
            textBox.LostFocus += new EventHandler(textBox_LostFocus);
            textBox.KeyDown += new KeyEventHandler(textBox_KeyDown);
        }

        private void textBox_LostFocus(object sender, EventArgs e)
        {
            Properties.Settings.Default[name] = textBox.Text;
        }

        protected void textBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                textBox.Parent.Focus();
            }
        }
        #endregion
        public override void LoadValue() { textBox.Text = ((int)Properties.Settings.Default[name]).ToString(); }
        public new int Value
        {
            get { return (int)Properties.Settings.Default[name]; }
        }
    }

    #endregion

    #endregion

    #endregion
}


