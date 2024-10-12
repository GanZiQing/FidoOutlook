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

namespace OutlookAutomation
{
    public partial class AdvanceExport : UserControl
    {
        #region Init
        public AdvanceExport()
        {
            InitializeComponent();
            //StartPosition = FormStartPosition.CenterScreen;
            SubscribeToEvents();
            LoadSettings();

            //listView.Columns
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

                projectTracker = ReadJsonToDictionary<string, Dictionary<string, Dictionary<string, string>>>(linkedJsonPath);
                
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
        public TabPage AdvanceExportTabPage
        {
            get
            {
                return advanceExportTabPage;
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
        /// [Key: Filter Name]   
        /// [Value: Dictionary 3]
        /// </para>
        /// <para>
        /// Dictionary 3
        /// [Key: Filter Parameter/Column Name e.g. Name, recipient, sender]
        /// [Value: Filter Parameter/Column Value]
        /// </para>
        /// </summary>
        public Dictionary<string, Dictionary<string, Dictionary<string, string>>> projectTracker = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();
        private void addNewProject_Click(object sender, EventArgs e)
        {
            CreateFilters createFilters = new CreateFilters(this);
            createFilters.FormClosed += new FormClosedEventHandler(AddProjectToTracker);
            createFilters.Show();

            // when createFilter closes, save value from createFilter
        }

        private void AddProjectToTracker(object sender, FormClosedEventArgs e)
        {
            CreateFilters createFilters = (CreateFilters)sender;
            if (!createFilters.setValue) { return; }

            if (projectTracker.ContainsKey(createFilters.projectName))
            {
                throw new ArgumentException($"Unable to create project as similar project name already exists.\nProject Name: {createFilters.projectName}");
            }

            projectTracker[createFilters.projectName] = createFilters.projectDictionary;
            RefreshListBox(true);
            
            MessageBox.Show($"New project {createFilters.projectName} added", "Project Added");
        }

        #endregion

        #region Edit Project
        Dictionary<string, CreateFilters> projectUnderEdit = new Dictionary<string, CreateFilters>();

        private void editProjectButton_Click(object sender, EventArgs e)
        {
            //if (listView.SelectedItems.Count != 1) { MessageBox.Show("Select exactly one item to edit", "Error"); return; }
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
                CreateFilters form = projectUnderEdit[projectName];

                form.WindowState = FormWindowState.Minimized;
                form.Show();
                form.WindowState = FormWindowState.Normal;
                return;
            }
            #endregion

            CreateFilters createFilters = new CreateFilters(this, projectName, projectTracker[projectName]);
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
            CreateFilters createFilters = (CreateFilters)sender;
            if (!createFilters.setValue) 
            {
                projectUnderEdit.Remove(createFilters.originalProjectName);
                return; 
            }

            //Save Project
            if (createFilters.originalProjectName != createFilters.projectName) 
            {
                // Project renamed, remove old reference
                projectTracker.Remove(createFilters.originalProjectName);
            }

            projectTracker[createFilters.projectName] = createFilters.projectDictionary;
            
            projectUnderEdit.Remove(createFilters.originalProjectName);

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

            projectTracker = ReadJsonToDictionary<string, Dictionary<string, Dictionary<string, string>>>(linkedJsonPath);
            RefreshListBox();
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
            projectTracker = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();
            RefreshListBox();

            MessageBox.Show($"Linked Json file removed","Completed");
        }
        private void importJson_Click(object sender, EventArgs e)
        {
            if (linkedJsonPath != "")
            {
                DialogResult result = MessageBox.Show("Json file currently linked. Append to current linked json?", "Warning", MessageBoxButtons.YesNo);
                if (result != DialogResult.Yes) { return; }
                //unlinkJson_Click(sender, e);
            }

            #region Get File Path
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "json files (*.json)|*.json";
            DialogResult res = dialog.ShowDialog();
            if (res == DialogResult.Cancel) { return; }
            //linkedJsonPath = dialog.FileName;
            #endregion

            #region Load Json and combine dictionary
            var importProjectTracker = ReadJsonToDictionary<string, Dictionary<string, Dictionary<string, string>>>(dialog.FileName);
            bool? overWrite = null;
            int numImport = 0;
            foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, string>>> entry in importProjectTracker)
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

        #region Export Buttons
        private void exportSelectedWithFilter_Click(object sender, EventArgs e)
        {
            sortedCriteria = new Dictionary<string, Dictionary<string, HashSet<string>>>();
            #region Create Filter Object
            foreach (ListViewItem item in listView.SelectedItems) // Loop through projects
            {
                string projectName = item.Text;
                AddToSortedFilter(projectName);
            }
            #endregion

            #region Export mailItems
            linkedPrintPane.exportWithFilter(sortedCriteria);
            #endregion
        }


        /// <summary> 
        /// <para>
        /// Outer dictionary 
        /// [Key: Filter 1 = recipient]   
        /// [Value : Filter 2 Dictionary]
        /// </para>
        /// <para>
        /// Inner dictionary
        /// [Key: Filter 2 = sender]   
        /// [Value: List of folderPath]
        /// </para>
        /// </summary>
        Dictionary<string, Dictionary<string, HashSet<string>>> sortedCriteria = new Dictionary<string, Dictionary<string, HashSet<string>>>();
        private void exportAllWithFilter_Click(object sender, EventArgs e)
        {
            sortedCriteria = new Dictionary<string, Dictionary<string, HashSet<string>>>();
            #region Create Filter Object
            foreach (ListViewItem item in listView.Items) // Loop through projects
            {
                string projectName = item.Text;
                AddToSortedFilter(projectName);
            }
            #endregion

            #region Export mailItems
            linkedPrintPane.exportWithFilter(sortedCriteria);
            #endregion
        }
        private void AddToSortedFilter(string projectName)
        {
            Dictionary<string, Dictionary<string, string>> projectFilters = projectTracker[projectName];

            foreach (KeyValuePair<string, Dictionary<string, string>> entry in projectFilters) // Loop through each filter type
            {
                string filterName = entry.Key;
                Dictionary<string,string> filterParameter = entry.Value;

                string sender = filterParameter["sender"];
                string recipient = filterParameter["recipient"];
                string folderPath = filterParameter["folderPath"];

                if (!sortedCriteria.ContainsKey(recipient))
                {
                    sortedCriteria[recipient] = new Dictionary<string, HashSet<string>>();
                }
                var recipientDictionary = sortedCriteria[recipient];


                if (!recipientDictionary.ContainsKey(sender))
                {
                    recipientDictionary[sender] = new HashSet<string>();
                }
                HashSet<string> senderList = recipientDictionary[sender];

                if (!senderList.Contains(recipient)) { senderList.Add(folderPath); }
                
            }
        }
        #endregion
    }

    partial class PrintPane
    {
        public void exportWithFilter(Dictionary<string, Dictionary<string, HashSet<string>>> sortedCriteria)
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
                    //progressTracker.UpdateStatus($"Getting mail items");
                    Application outlookApp = Globals.ThisAddIn.Application;
                    Explorer explorer = outlookApp.ActiveExplorer();
                    maxItems = explorer.Selection.Count;
                    if (maxItems == 0) { MessageBox.Show("No items selected."); return; }
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
                            thisCustomMailItem = new CustomMailItem(mailItem, exportOptions, sortedCriteria, wordApp);
                            progressTracker.UpdateStatus($"Exporting: {mailItem.Subject}");

                            thisCustomMailItem.ExportWithFilters();

                            #endregion

                            #region Increment and Update Progress
                            ////worker.ReportProgress(GlobalUtilities.ConvertToProgress(currentIndex, maxItems));
                            //if (worker.CancellationPending)
                            //{
                            //    return;
                            //}
                            currentIndex += 1;
                            #endregion

                        }
                        #region Catch Finally
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
                            if (moveItemsCheck.Checked)
                            {
                                thisCustomMailItem.MoveToFolder();
                            }

                            if (mailItem != null) { Marshal.ReleaseComObject(mailItem); }
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
                if (thisCustomMailItem != null) { thisCustomMailItem.ReleaseItems(); }
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.FinalReleaseComObject(wordApp);
                    wordApp = null;
                }
                Beaver.CheckLog();
            }
        }
    }
}
