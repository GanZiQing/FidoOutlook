using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static OutlookAutomation.ExportUtilities;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using OutlookAutomation.Properties;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Microsoft.Office.Interop.Outlook;
using static OutlookAutomation.OutlookUtilities;
using Exception = System.Exception;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;
using Newtonsoft.Json.Linq;

namespace OutlookAutomation
{
    public partial class HdbExport : UserControl
    {
        #region Init
        public HdbExport()
        {
            InitializeComponent();
            SubscribeToEvents();
            LoadSettings();
        }

        private void SubscribeToEvents()
        {
            listView.MouseDoubleClick += new MouseEventHandler(listView_MouseDoubleClick);
            listView.ColumnClick += new ColumnClickEventHandler(listView_ColumnClick);
        }

        private void listView_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (listView.Sorting == SortOrder.Ascending)
            {
                listView.Sorting = SortOrder.Descending;
            }
            else if (listView.Sorting == SortOrder.Descending)
            {
                listView.Sorting = SortOrder.None;
                RefreshListBox();
            }
            else if (listView.Sorting == SortOrder.None)
            {
                listView.Sorting = SortOrder.Ascending;
            }
        }

        private void listView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ListViewHitTestInfo info = listView.HitTest(e.X, e.Y);
                System.Windows.Forms.ListViewItem item = info.Item;

                if (item == null) { return; }

                string projectName = item.Text;
                LaunchEditItem(projectName);
            }
        }

        private void LoadSettings()
        {
            linkedJsonPath = Settings.Default.LinkedJsonFile;
            if (linkedJsonPath == null) { linkedJsonPath = ""; }
            dispLinkedPath.Text = linkedJsonPath;

            if (linkedJsonPath == "" || linkedJsonPath == null) // No file set
            {
                return;
            }

            try
            {
                if (!File.Exists(linkedJsonPath))
                {
                    throw new ArgumentException($"Unable to find Json file");
                }
                
                projectTracker = ReadJsonForHDBExport(linkedJsonPath);
                
                RefreshListBox();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error importing json file at {linkedJsonPath}\n{ex.Message}");
            }
        }

        string folderPath = "";
        private void setExportPath_Click(object sender, EventArgs e)
        {
            CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
            DialogResult res = customFolderBrowser.ShowDialog();
            if (res == DialogResult.Cancel) { return; }
            folderPath = customFolderBrowser.folderPath;
            MessageBox.Show($"Folder set to {folderPath}");
        }

        #region Link to Export Pane
        public TabPage HdbExportTab
        {
            get
            {
                return hdbExportTabPage;
            }
        }
        PrintPane linkedPrintPane;
        public PrintPane LinkedPrintPane
        {
            get
            {
                return linkedPrintPane;
            }
            set
            {
                linkedPrintPane = value;
            }
        }
        #endregion
        #endregion

        #region Add Project
        /// <summary> 
        /// <para>
        /// Dictionary 1
        /// [Key: Project Name]   
        /// [Value : Dictionary 2]
        /// </para>
        /// <para>
        /// Dictionary 2
        /// [Key: Table Name]   
        /// [Value: Object 3]
        /// </para>
        /// <para>
        /// Object 3 - may be a dictionary (recipient/sender) or hashset (internal sender)
        /// [Key: Email/SubjectText]
        /// [Value: ReplacementText/Number of Char.]
        /// </para>
        /// </summary>
        public Dictionary<string, Dictionary<string, object>> projectTracker = new Dictionary<string, Dictionary<string, object>>();
        private void addNewProject_Click(object sender, EventArgs e)
        {
            HdbFilters createFilters = new HdbFilters(this);
            createFilters.FormClosed += new FormClosedEventHandler(AddProjectToTracker);
            createFilters.Show(this);

            // when createFilter closes, save value from createFilter
        }

        private void AddProjectToTracker(object sender, FormClosedEventArgs e)
        {
            HdbFilters filters = (HdbFilters)sender;
            if (!filters.setValue) { return; }

            if (projectTracker.ContainsKey(filters.projectName))
            {
                throw new ArgumentException($"Unable to create project as similar project name already exists.\nProject Name: {filters.projectName}");
            }

            projectTracker[filters.projectName] = filters.projectDictionary;
            RefreshListBox(true);
            
            MessageBox.Show($"New project {filters.projectName} added", "Project Added");
        }

        #endregion

        #region Edit Project
        Dictionary<string, HdbFilters> projectUnderEdit = new Dictionary<string, HdbFilters>();

        private void editProjectButton_Click(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count == 0) { MessageBox.Show("Select item(s) to edit", "Error"); return; }
            foreach (var item in listView.SelectedItems)
            {
                string projectName = listView.SelectedItems[0].Text;
                LaunchEditItem(projectName);
            }
        }

        private void LaunchEditItem(string projectName)
        {
            #region Get Existing Open Form if Available
            if (projectUnderEdit.ContainsKey(projectName))
            {
                HdbFilters form = projectUnderEdit[projectName];

                form.WindowState = FormWindowState.Minimized;
                form.Show();
                form.WindowState = FormWindowState.Normal;
                return;
            }
            #endregion

            HdbFilters createFilters = new HdbFilters(this, projectName, projectTracker[projectName]);
            createFilters.FormClosed += new FormClosedEventHandler(SaveEditedProjectToTracker);
            projectUnderEdit[projectName] = createFilters;
            createFilters.Show();
        }

        private void deleteProjectButton_Click(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count == 0) { MessageBox.Show("No item selected to delete", "Error"); return; }

            DialogResult result = MessageBox.Show($"Delete selected project?\n{listView.SelectedItems.Count} projects selected.", "Delete project", MessageBoxButtons.YesNo);
            if (result != DialogResult.Yes) { return; }

            foreach (ListViewItem item in listView.SelectedItems)
            {
                string projectName = item.Text;
                projectTracker.Remove(projectName);
            }
            RefreshListBox(true);
        }
        private void SaveEditedProjectToTracker(object sender, FormClosedEventArgs e)
        {
            HdbFilters hdbFilters = (HdbFilters)sender;
            if (!hdbFilters.setValue) 
            {
                projectUnderEdit.Remove(hdbFilters.originalProjectName);
                return; 
            }

            //Save Project
            if (hdbFilters.originalProjectName != hdbFilters.projectName) 
            {
                // Project renamed, remove old reference
                projectTracker.Remove(hdbFilters.originalProjectName);
            }

            projectTracker[hdbFilters.projectName] = hdbFilters.projectDictionary;
            
            projectUnderEdit.Remove(hdbFilters.originalProjectName);

            RefreshListBox(true);
        }
        #endregion

        #region Json
        #region Importing Json
        string linkedJsonPath;
        private void linkJsonFile_Click(object sender, EventArgs e)
        {
            #region Warning
            DialogResult result = MessageBox.Show("Linking json file will unlink the current json (if any). " +
                "If no Json is linked, current project details will be lost.", "Confirmation", MessageBoxButtons.YesNo);
            if (result != DialogResult.Yes) { return; }
            #endregion
            
            #region Get File Path
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "json files (*.json)|*.json";
            DialogResult res = dialog.ShowDialog();
            if (res == DialogResult.Cancel) { return; }
            linkedJsonPath = dialog.FileName;
            #endregion

            Settings.Default.LinkedJsonFile = linkedJsonPath;
            Settings.Default.Save();
            dispLinkedPath.Text = linkedJsonPath;

            projectTracker = ReadJsonForHDBExport(linkedJsonPath);
            RefreshListBox();
        }

        public static Dictionary<string, Dictionary<string, object>> ReadJsonForHDBExport(string filePath)
        {
            var readDict = ReadJsonToObject<Dictionary<string, object>>(filePath);
            Dictionary<string, Dictionary<string, object>> projectTrackingImport = new Dictionary<string, Dictionary<string, object>>();
            foreach (var entry in readDict)
            {
                string key = entry.Key;
                JsonElement value = (JsonElement)entry.Value;
                Dictionary<string, object> projectDictionary = DecomposeEachProjectDictionary(value);
                projectTrackingImport[key] = projectDictionary;
            }

            return projectTrackingImport;
        }
        private static Dictionary<string, object> DecomposeEachProjectDictionary(JsonElement obj)
        {
            var readDict = obj.Deserialize<Dictionary<string, object>>();
            Dictionary<string, object> projectDictionary = new Dictionary<string, object>();
            foreach (var entry in readDict)
            {
                string key = entry.Key;
                JsonElement value = (JsonElement)entry.Value;
                if (value.ValueKind == JsonValueKind.Array)
                {
                    HashSet<string> hashSet = value.Deserialize<HashSet<string>>();
                    projectDictionary.Add(key, hashSet);
                }
                else if (value.ValueKind == JsonValueKind.Object)
                {
                    Dictionary<string, string> dict = value.Deserialize<Dictionary<string, string>>();
                    projectDictionary.Add(key, dict);
                }
                else
                {
                    throw new ArgumentException("Unable to decompose Json value");
                }
            }

            return projectDictionary;
        }
        private void unlinkJson_Click(object sender, EventArgs e)
        {
            if (linkedJsonPath == "") 
            {
                MessageBox.Show("No Json file linked");
            }

            linkedJsonPath = "";
            Settings.Default.LinkedJsonFile = linkedJsonPath;
            Settings.Default.Save();

            dispLinkedPath.Text = linkedJsonPath;
            
            projectTracker = new Dictionary<string, Dictionary<string, object>>();
            
            RefreshListBox();

            MessageBox.Show($"Linked Json file removed","Completed");
        }
        private void importJson_Click(object sender, EventArgs e)
        {
            if (linkedJsonPath != "")
            {
                DialogResult result = MessageBox.Show("Json file currently linked. Append to current linked json?", "Warning", MessageBoxButtons.YesNo);
                if (result != DialogResult.Yes) { return; }
            }

            #region Get File Path
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "json files (*.json)|*.json";
            DialogResult res = dialog.ShowDialog();
            if (res == DialogResult.Cancel) { return; }
            //linkedJsonPath = dialog.FileName;
            #endregion

            #region Load Json and combine dictionary
            var importProjectTracker = ReadJsonForHDBExport(dialog.FileName);
            bool? overWrite = null;
            int numImport = 0;
            foreach (KeyValuePair<string, Dictionary<string, object>> entry in importProjectTracker)
            {
                string importProjectName = entry.Key;
                if (projectTracker.ContainsKey(importProjectName))
                {
                    if (overWrite == null)
                    {
                        DialogResult result = MessageBox.Show($"Duplicate project name {importProjectName} found. \n" +
                        $"Overwrite duplicate projects? This will apply for all subsequent clashes.\n" +
                        $"Cancel will terminate import.", "Warning", MessageBoxButtons.YesNoCancel);
                        if (result == DialogResult.Cancel) { return; }
                        else if (result == DialogResult.Yes) { overWrite = true; }
                        else { overWrite = false; }
                    }

                    if (!(bool)overWrite)
                    {
                        continue;
                    }
                }
                
                projectTracker[importProjectName] = importProjectTracker[importProjectName];
                numImport++;
            }
            #endregion

            RefreshListBox(true);
            MessageBox.Show($"{numImport}/{importProjectTracker.Count} projects succesfully imported");
        }

        #endregion
        
        #region Exporting Json
        private void exportJson_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "json files (*.json)|*.json";
                DialogResult result = saveFileDialog.ShowDialog();
                if (result == DialogResult.Cancel) { return; }
                string savePath = saveFileDialog.FileName;

                WriteToJson(projectTracker, savePath);
                MessageBox.Show($"Data saved to linked json file\n" +
                    $"{savePath}", "Completed");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to export json file to {linkedJsonPath}.\n" +
                    $"{ex.Message}");
            }
        }

        private void SaveLinkedJson()
        {
            try
            {
                if (linkedJsonPath == null || linkedJsonPath == "") { return; }
                string savePath = linkedJsonPath;
                if (!File.Exists(savePath))
                {
                    DialogResult result = MessageBox.Show($"Linked Json at {linkedJsonPath} does not exist. Continue? New file will be created", "Error", MessageBoxButtons.YesNo);
                    if (result != DialogResult.Yes)
                    {
                        DialogResult unlinkResult = MessageBox.Show($"Unlink Json File?", "Error", MessageBoxButtons.YesNo);
                        if (unlinkResult == DialogResult.Yes)
                        {
                            linkedJsonPath = "";
                            Settings.Default.LinkedJsonFile = linkedJsonPath;
                            Settings.Default.Save();

                            dispLinkedPath.Text = linkedJsonPath;
                        }
                        return;
                    }
                }
                WriteToJson(projectTracker, savePath);
                string dateString = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");
                lastSavedLabel.Text = $"Last Saved: {dateString}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to export json file to {linkedJsonPath}.\n" +
                    $"{ex.Message}");
            }
        }
        #endregion
        #endregion

        #region Update ListBox
        private void RefreshListBox(bool saveToLinked = false)
        {
            listView.Items.Clear();

            foreach (string item in projectTracker.Keys)
            {
                listView.Items.Add(item);
            }
            if (saveToLinked)
            {
                SaveLinkedJson();
            }
        }
        #endregion

        private void exportSelectedOnly_Click(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count == 0) { MessageBox.Show("No project selected.", "Error"); return; }
            DialogResult dialogResult = MessageBox.Show($"Export for project: {listView.SelectedItems[0].Text}", "Confirmation", MessageBoxButtons.YesNo);
            if (dialogResult != DialogResult.Yes) { return; }
            Dictionary<string, object> hdbExportCriteria = projectTracker[listView.SelectedItems[0].Text];
            linkedPrintPane.ExportForHDB(hdbExportCriteria);
        }
    }

    partial class PrintPane
    {
        public void ExportForHDB(Dictionary<string, object> hdbExportCriteria)
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
                exportOptions.SaveSettings();
                Beaver.logExist = false;
                Beaver.Initialize(exportOptions.baseFolder, "Export Error Log.txt");

                if (exportOptions.pdf)
                {
                    wordApp = new Word.Application();
                    wordApp.Visible = true; //Must be visible because of security notifications
                    wordApp.Activate();
                }

                #endregion

                #region Debugging Zone
                //Application outlookAppTest = Globals.ThisAddIn.Application;
                //MailItem mailItemTest = outlookAppTest.Session.OpenSharedItem(filePaths[0]);
                //var recept = mailItemTest.Recipients;
                //string referenceName0 = mailItemTest.Recipients[1].Name;
                //string referenceName = mailItemTest.Recipients[0].Name;
                #endregion

                ProgressHelper.RunWithProgress((worker, progressTracker) =>
                {
                    #region Explorer
                    progressTracker.UpdateStatus($"Getting mail items");
                    Application outlookApp = Globals.ThisAddIn.Application;
                    //Explorer explorer = outlookApp.ActiveExplorer();
                    maxItems = filePaths.Length;
                    #endregion

                    foreach (string filePath in filePaths)
                    {
                        MailItem mailItem = outlookApp.Session.OpenSharedItem(filePath);
                        try
                        {
                            #region Export
                            thisCustomMailItem = new CustomMailItem(mailItem, filePath, exportOptions, hdbExportCriteria, wordApp);
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
                        #region Catch Finally
                        catch (Exception ex)
                        {
                            string msg = $"Unable complete export mail function.\n" +
                            $"    Item: {currentIndex}/{maxItems}\n" +
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
                            if (mailItem != null) { mailItem.Close(OlInspectorClose.olDiscard); Marshal.ReleaseComObject(mailItem); mailItem = null; }
                            if (thisCustomMailItem != null) { thisCustomMailItem.ReleaseItems(); }
                        }
                        #endregion
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
    }

}
