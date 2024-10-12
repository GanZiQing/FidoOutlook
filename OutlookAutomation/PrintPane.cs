using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Outlook.Application;
using System.IO;
using System.Runtime.InteropServices;
using Exception = System.Exception;

namespace OutlookAutomation
{
    public partial class PrintPane : UserControl
    {
        #region Init
        ExportOptions exportOptions;
        public PrintPane()
        {
            InitializeComponent();
            AddToolTips();
            AddTabPage();
            InitialiseExportOptions();
        }

        //AdvanceExport advanceExport;
        HdbExport advanceExport;
        private void AddTabPage()
        {
            //summonForm = new AdvanceExport();
            advanceExport = new HdbExport();
            advanceExport.LinkedPrintPane = this;
            //tabControl1.TabPages.Insert(0, summonForm.AdvanceExportTabPage);
            tabControl1.TabPages.Add(advanceExport.HdbExportTab);
            TabPage tabPage = tabControl1.TabPages[0];
            tabControl1.TabPages[0] = tabControl1.TabPages[1];
            tabControl1.TabPages[1] = tabPage;
        }

        private void AddToolTips()
        {
            toolTip1.SetToolTip(exportPdfCheck, "Warning: Experimental function. Not a native outlook feature, more computationally intensive. Check output.\n" +
                "Opens an instance of word (.rtf) to save pdf. May require user to allow access to rtf for internet emails.");
            toolTip1.SetToolTip(saveSettings, "Also automatically saved upon export.");
            toolTip1.SetToolTip(skipEmbeddedCheck, "Warning: Experimental function. Use to skip embeded images like signature logos. Check output. \nDisable if it is causing issues");
        }
        #endregion

        #region Export Options
        private void InitialiseExportOptions()
        {
            exportOptions = new ExportOptions();
            exportOptions.Add(new CheckSetting("exportAttachments", exportAttachmentsCheck));
            exportOptions.Add(new CheckSetting("exportHtml", exportHtmlCheck));
            exportOptions.Add(new CheckSetting("exportMsg", exportMsgCheck));
            exportOptions.Add(new CheckSetting("exportPdf", exportPdfCheck));
            exportOptions.Add(new CheckSetting("exportWord", exportWordCheck));

            exportOptions.Add(new CheckSetting("dateFolder", dateFolderCheck));
            exportOptions.Add(new CheckSetting("subjectFolder", subjectFolderCheck));

            exportOptions.Add(new CheckSetting("skipEmbedded", skipEmbeddedCheck));
            exportOptions.Add(new CheckSetting("shortenSubject", shortenSubjectCheck));
            exportOptions.Add(new IntBoxSetting("maxSubjectLength", dispMaxSubjectLength));
            exportOptions.Add(new CheckSetting("breakOnError", breakOnErrorCheck));
            exportOptions.Add(new CheckSetting("moveItems", moveItemsCheck));

            exportOptions.Add(new TextBoxSetting("baseFolder", dispBaseFolder));
            exportOptions.LoadSettings();
            //exportAttachmentsCheck.Checked = exportOptions.attachments;
            //exportHtmlCheck.Checked = exportOptions.html;
            //exportMsgCheck.Checked = exportOptions.msg;
            //exportPdfCheck.Checked = exportOptions.pdf;
            //exportWordCheck.Checked = exportOptions.word;

            //subjectFolderCheck.Checked = exportOptions.subjectFolder;
            //dateFolderCheck.Checked = exportOptions.dateTimeFolder;

            //skipEmbededCheck.Checked = exportOptions.skipEmbedded;
            //shortenSubjectCheck.Checked = exportOptions.shortenSubject;
            //maxSubjectLength.Text = exportOptions.maxSubjectLength.ToString();

            //dispBaseFolder.Text = exportOptions.baseFolder;
        }

        private void saveSettingsButton_Click(object sender, EventArgs e)
        {
            try
            {
                //GetExportOptions();
                //exportOptions.SaveToSettings();
                exportOptions.SaveSettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        //private void GetExportOptions()
        //{
        //    // Rewrite export options to track the checkbox object instead
        //    exportOptions.RefreshFromUserControl(
        //        exportAttachmentsCheck.Checked,
        //        exportHtmlCheck.Checked,
        //        exportMsgCheck.Checked,
        //        exportPdfCheck.Checked,
        //        exportWordCheck.Checked,

        //        subjectFolderCheck.Checked,
        //        dateFolderCheck.Checked,

        //        skipEmbeddedCheck.Checked,
        //        shortenSubjectCheck.Checked,
        //        dispMaxSubjectLength.Text,
        //        breakOnErrorCheck.Checked,
        //        dispBaseFolder.Text,
        //        moveItemsCheck.Checked
        //        );
        //}
        #endregion

        #region Export Operation
        List<MailItem> failedMailItem = new List<MailItem>();
        private void exportAllSelected_Click(object sender, EventArgs e)
        {
            CustomMailItem thisCustomMailItem = null;
            Word.Application wordApp = null;
            int maxItems = 0;
            int currentIndex = 0;
            try
            {
                #region Initialise
                //GetExportOptions();
                Beaver.logExist = false;
                Beaver.Initialize(exportOptions.baseFolder, "Export Error Log.txt");

                if (exportOptions.pdf)
                {
                    wordApp = new Word.Application();
                    wordApp.Visible = true; //Debug only
                }
                
                #endregion
                
                ProgressHelper.RunWithProgress((worker, progressTracker) =>
                {
                    #region Explorer
                    progressTracker.UpdateStatus($"Getting mail items");
                    Application outlookApp = Globals.ThisAddIn.Application;
                    Explorer explorer = outlookApp.ActiveExplorer();
                    maxItems = explorer.Selection.Count;
                    #endregion

                    foreach (object item in explorer.Selection)
                    {
                        MailItem mailItem = null;
                        try
                        {
                            #region Ensure item is MailItem
                            
                            if (item is MailItem)
                            {
                                mailItem = (MailItem)item;
                            }
                            else
                            {
                                Beaver.LogError($"Selected item isn't a mail item, skipped.");
                                continue;
                            }
                            #endregion

                            #region Export
                            thisCustomMailItem = new CustomMailItem(mailItem, exportOptions, wordApp);
                            progressTracker.UpdateStatus($"Exporting: {mailItem.Subject}");
                            thisCustomMailItem.Export();
                            #endregion

                            #region Increment and Update Progress
                            worker.ReportProgress(GlobalUtilities.ConvertToProgress(currentIndex, maxItems));
                            if (worker.CancellationPending)
                            {
                                return;
                            }
                            currentIndex += 1;
                            #endregion
                            
                        }
                        catch (Exception ex)
                        {
                            string msg = $"Unable complete export mail function.\n" +
                            $"    Item: {currentIndex}/{maxItems}" +
                            $"    Subject: {mailItem.Subject}\n" +
                            $"    Date: {mailItem.SentOn.ToString("dddd, dd MMMM yyyy h:mm tt")}\n" +
                            $"    Error Message: {ex.Message}\n";

                            Beaver.LogError(msg);

                            failedMailItem.Add(mailItem);
                            currentIndex += 1;

                            if (exportOptions.breakOnError)
                            {
                                throw new Exception("Error encountered in export, terminating");
                            }
                            else
                            {
                                continue;
                            }
                        }
                        finally
                        {
                            if (mailItem != null) { Marshal.ReleaseComObject(mailItem); }
                            if (thisCustomMailItem != null) { thisCustomMailItem.ReleaseItems(); }
                        }
                    }
                    progressTracker.UpdateStatus($"Completed, check message box.");
                    MessageBox.Show("Completed", "Completed");
                });
                
            }
            catch (Exception ex)
            {
                if (currentIndex != maxItems)
                {
                    Beaver.LogError($"Terminated at item {currentIndex}/{maxItems}\n");
                }
                MessageBox.Show(ex.Message, "Error");
                
            }
            finally
            {
                if (thisCustomMailItem != null) { thisCustomMailItem.ReleaseItems(); }
                if (wordApp!= null)
                {
                    wordApp.Quit();
                    Marshal.FinalReleaseComObject(wordApp);
                    wordApp = null;
                }
                Beaver.CheckLog();
            }
        }
        #endregion

        #region Helper Functions
        private MailItem GetCurrentMailItem()
        {
            Application outlookApp = Globals.ThisAddIn.Application;
            Explorer explorer = outlookApp.ActiveExplorer();

            #region Checks
            if (explorer == null)
            {
                throw new Exception("No explorer found");
            }
            // Get the current item in the Reading Pane
            if (explorer.Selection.Count == 0)
            {
                throw new Exception("No item selected");
            }

            if (!(explorer.Selection[1] is MailItem))
            {
                throw new Exception("No mail item selected");
            }
            #endregion

            return explorer.Selection[1];
        }

        private List<MailItem> GetCurrentMailItems()
        {
            Application outlookApp = Globals.ThisAddIn.Application;
            Explorer explorer = outlookApp.ActiveExplorer();

            #region Checks
            if (explorer == null)
            {
                throw new Exception("No explorer found");
            }
            // Get the current item in the Reading Pane
            if (explorer.Selection.Count == 0)
            {
                throw new Exception("No item selected");
            }

            #endregion

            List<MailItem> allMailItems = new List<MailItem>();
            foreach (object item in explorer.Selection)
            {
                if (item is MailItem)
                {
                    allMailItems.Add((MailItem)item);
                }
                else
                {
                    Beaver.LogError($"Selected item isn't a mail item, skipped.");
                }
            }
            if (allMailItems.Count == 0)
            {
                throw new System.Exception("No mail item selected");
            }
            return allMailItems;
        }
        

        #endregion

        #region Directory
        private void setFolder_Click(object sender, EventArgs e)
        {
            CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
            if (customFolderBrowser.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            string baseFolder = customFolderBrowser.GetFolderPath();
            dispBaseFolder.Text = baseFolder;
        }

        private void openFolder_Click(object sender, EventArgs e)
        {
            try
            {
                // Get Path
                string folderPath = dispBaseFolder.Text;

                //Check if path exist
                if (folderPath == "")
                {
                    throw new ArgumentException("No path provided");
                }

                if (!Directory.Exists(folderPath))
                {
                    throw new ArgumentException($"Directory provided is invalid: \n{folderPath}");
                }
                else
                {
                    System.Diagnostics.Process.Start(folderPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        #endregion

        #region Others
        private void selectFailed_Click(object sender, EventArgs e)
        {
            try
            {
                if (failedMailItem.Count == 0) { MessageBox.Show("No mailitem to select", "Error"); }
                Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                explorer.ClearSelection();
                int failedToSelect = 0;
                foreach (MailItem mailItem in failedMailItem)
                {
                    try
                    {
                        explorer.AddToSelection(mailItem);
                    }
                    catch
                    {
                        failedToSelect++;
                    }

                }
                if (failedToSelect > 0)
                {
                    MessageBox.Show($"Warning: Unable to select {failedToSelect} items", "Warning");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        #endregion

        #region Export .msg
        private void exportMSG_Click(object sender, EventArgs e)
        {
            #region Open file
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "MSG (.msg)|*.msg";
            dialog.Multiselect = true;
            DialogResult dialogResult = dialog.ShowDialog();
            string[] filePaths;
            if (dialogResult == DialogResult.OK)
            {
                filePaths = dialog.FileNames;
            }
            else
            {
                return;
            }
            #endregion

            #region Export As MailItem

            CustomMailItem thisCustomMailItem = null;
            Word.Application wordApp = null;
            int maxItems = 0;
            int currentIndex = 0;
            try
            {
                #region Initialise
                //GetExportOptions();
                Beaver.logExist = false;
                Beaver.Initialize(exportOptions.baseFolder, "Export Error Log.txt");

                if (exportOptions.pdf)
                {
                    wordApp = new Word.Application();
                    wordApp.Visible = true; //Debug only
                }

                #endregion

                ProgressHelper.RunWithProgress((worker, progressTracker) =>
                {
                    #region Explorer
                    progressTracker.UpdateStatus($"Getting mail items");
                    Application outlookApp = Globals.ThisAddIn.Application;
                    Explorer explorer = outlookApp.ActiveExplorer();
                    maxItems = filePaths.Length;
                    #endregion

                    foreach (string filePath in filePaths)
                    {
                        MailItem mailItem = outlookApp.Session.OpenSharedItem(filePath);
                        try
                        {
                            #region Export
                            thisCustomMailItem = new CustomMailItem(mailItem, exportOptions, wordApp);
                            progressTracker.UpdateStatus($"Exporting: {mailItem.Subject}");
                            thisCustomMailItem.Export();
                            #endregion

                            #region Increment and Update Progress
                            worker.ReportProgress(GlobalUtilities.ConvertToProgress(currentIndex, maxItems));
                            if (worker.CancellationPending)
                            {
                                return;
                            }
                            currentIndex += 1;
                            #endregion

                        }
                        catch (Exception ex)
                        {
                            string msg = $"Unable complete export mail function.\n" +
                            $"    Item: {currentIndex}/{maxItems}" +
                            $"    Subject: {mailItem.Subject}\n" +
                            $"    Date: {mailItem.SentOn.ToString("dddd, dd MMMM yyyy h:mm tt")}\n" +
                            $"    Error Message: {ex.Message}\n";

                            Beaver.LogError(msg);

                            failedMailItem.Add(mailItem);
                            currentIndex += 1;

                            if (exportOptions.breakOnError)
                            {
                                throw new Exception("Error encountered in export, terminating");
                            }
                            else
                            {
                                continue;
                            }
                        }
                        finally
                        {
                            if (mailItem != null) { mailItem.Close(OlInspectorClose.olDiscard); Marshal.ReleaseComObject(mailItem); mailItem = null;  }
                            if (thisCustomMailItem != null) { thisCustomMailItem.ReleaseItems(); }
                            
                        }
                    }
                    progressTracker.UpdateStatus($"Completed, check message box.");
                    MessageBox.Show("Completed", "Completed");
                });

            }
            catch (Exception ex)
            {
                if (currentIndex != maxItems)
                {
                    Beaver.LogError($"Terminated at item {currentIndex}/{maxItems}\n");
                }
                MessageBox.Show(ex.Message, "Error");

            }
            finally
            {
                if (thisCustomMailItem != null) { thisCustomMailItem.ReleaseItems(); thisCustomMailItem = null; }
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.FinalReleaseComObject(wordApp);
                    wordApp = null;
                }
                Beaver.CheckLog();
                GC.Collect();
            }
            #endregion
        }
        #endregion
    }
}



