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
using System.Runtime.Remoting.Contexts;
using System.Text.Json;

namespace OutlookAutomation
{
    public partial class HdbFilters : Form
    {
        #region Init
        public bool isNew = false;
        public HdbFilters(HdbExport parentForm) 
        {
            InitializeComponent();
            CancelButton = cancelButton;
            StartPosition = FormStartPosition.CenterScreen;
            this.parentForm = parentForm;
            isNew = true;
            Text = "Add New Project";
            SubscribeToEvents();
        }

        public string originalProjectName;
        public HdbFilters(HdbExport parentForm, string projectName, Dictionary<string, object> tableData)
        {
            InitializeComponent();
            dispProjectName.Text = projectName;
            originalProjectName = projectName;
            Text = originalProjectName;
            LoadDictionaryForHDBFilters(tableData);
        }

        #region Events
        private void SubscribeToEvents()
        {
            senderGridView.LostFocus += new EventHandler(DataGridLoseFocus);
            recipientGridView.LostFocus += new EventHandler(DataGridLoseFocus);
            subjectGridView.LostFocus += new EventHandler(DataGridLoseFocus);
        }
        private DataGridView lastGridView;
        private void DataGridLoseFocus(object sender, EventArgs e)
        {
            lastGridView = (DataGridView)sender;
        }
        #endregion
        #endregion

        #region Data Validation
        #region Main
        private bool CheckTables()
        {
            #region Check if all tables are empty
            if (CheckIfTableIsEmpty(senderGridView) && CheckIfTableIsEmpty(recipientGridView) && CheckIfTableIsEmpty(subjectGridView))
            {
                MessageBox.Show("All tables are empty", "Error");
                return false;
            }
            #endregion

            bool passCheck = true;           

            #region Check Each Table
            bool check = CheckRecipientTable();
            if (!check) { passCheck = false; }
            
            check = CheckSenderTable();
            if (!check) { passCheck = false; }

            check = CheckSubjectTable();
            if (!check) { passCheck = false; }
            #endregion

            return passCheck;
        }
        
        private bool CheckSenderTable()
        {
            DataGridView checkingTable = senderGridView;
            bool passCheck = true;
            checkingTable.ClearSelection();
            //#region Check Empty
            //bool isEmpty = CheckIfTableIsEmpty(checkingTable);
            //if (isEmpty)
            //{
            //    passCheck = false;
            //    MessageBox.Show($"{GetGridViewName(checkingTable)} Table is empty", "Error");
            //    return (!passCheck);
            //}
            //#endregion

            ResetTableColor(checkingTable);
            #region Check Rows
            if (!HighlightEmptyRows(checkingTable)) { passCheck = false; }
            if (!HighlightRepeatedRows(checkingTable, 0)) { passCheck = false; };
            #endregion

            #region Check Values
            for (int rowNum = 0; rowNum < checkingTable.RowCount - 1; rowNum++)
            {
                bool check = CheckIsNotEmpty(rowNum, 0, checkingTable);
                if (!check) { passCheck = false; }
            }
            #endregion

            return passCheck;
        }

        private bool CheckRecipientTable()
        {
            DataGridView checkingTable = recipientGridView;
            bool passCheck = true;
            checkingTable.ClearSelection();
            //#region Check Empty
            //bool isEmpty = CheckIfTableIsEmpty(checkingTable);
            //if (isEmpty)
            //{
            //    passCheck = false;
            //    MessageBox.Show($"{GetGridViewName(checkingTable)} Table is empty", "Error");
            //    return (!passCheck);
            //}
            //#endregion

            ResetTableColor(checkingTable);
            #region Check Rows
            if (!HighlightEmptyRows(checkingTable)) { passCheck = false; }
            if (!HighlightRepeatedRows(checkingTable, 0)) { passCheck = false; };
            #endregion

            #region Check Values
            for (int rowNum = 0; rowNum < checkingTable.RowCount - 1; rowNum++)
            {
                bool check = CheckIsNotEmpty(rowNum, 0, checkingTable);
                if (!check) { passCheck = false; }

                check = CheckIsNotEmpty(rowNum, 1, checkingTable);
                if (!check) { passCheck = false; }
            }
            #endregion

            return passCheck;
        }

        private bool CheckSubjectTable()
        {
            DataGridView checkingTable = subjectGridView;
            bool passCheck = true;
            checkingTable.ClearSelection();

            ResetTableColor(checkingTable);
            #region Check Rows
            if (!HighlightEmptyRows(checkingTable)) { passCheck = false; }
            if (!HighlightRepeatedRows(checkingTable, 1)) { passCheck = false; };
            #endregion

            #region Check Values
            for (int rowNum = 0; rowNum < checkingTable.RowCount - 1; rowNum++)
            {
                bool check = CheckIsNotEmpty(rowNum, 0, checkingTable);
                if (!check) { passCheck = false; }

                check = CheckIsNumber(rowNum, 1, checkingTable);
                if (!check) { passCheck = false; }
            }
            #endregion

            return passCheck;
        }
        #endregion

        #region Supporting Methods
        private void ResetTableColor()
        {
            foreach (DataGridViewRow row in recipientGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell.Style.BackColor = recipientGridView.DefaultCellStyle.BackColor;
                }
            }
        }
        private static bool CheckIfTableIsEmpty(DataGridView checkingTable)
        {
            for (int rowNum = 0; rowNum < checkingTable.RowCount - 1; rowNum++)
            {
                if (!CheckIfRowIsEmpty(checkingTable, rowNum))
                {
                    return false;
                }
            }
            return true;
        }

        private static bool CheckIfRowIsEmpty(DataGridView checkingTable, int rowNum)
        {
            for (int colNum = 0; colNum < checkingTable.ColumnCount; colNum++)
            {
                string value = (string)checkingTable.Rows[rowNum].Cells[colNum].Value;
                if (!(value == "" || value == null))
                {
                    return false;
                }
            }
            return true;
        }

        private static bool HighlightEmptyRows(DataGridView checkingTable)
        {
            List<int> emptyRowNums = new List<int>();
            for (int rowNum = 0; rowNum < checkingTable.RowCount - 1; rowNum++)
            {
                if (CheckIfRowIsEmpty(checkingTable, rowNum))
                {
                    emptyRowNums.Add(rowNum);
                }
            }

            if (emptyRowNums.Count == 0) { return true; }

            DialogResult res = MessageBox.Show("Empty rows found, delete rows?", "Empty Rows", MessageBoxButtons.YesNo);
            emptyRowNums.Reverse();
            if (res == DialogResult.Yes)
            {
                foreach (int rowNum in emptyRowNums)
                {
                    checkingTable.Rows.RemoveAt(rowNum);
                }
            }
            else
            {
                foreach (int rowNum in emptyRowNums)
                {
                    for (int colNum = 0; colNum < checkingTable.ColumnCount; colNum++)
                    {
                        checkingTable.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightGray;
                    }
                }
            }
            return false;
        }
        
        private static bool HighlightRepeatedRows(DataGridView checkingTable, int colNum)
        {
            Dictionary<string, int> existingValues = new Dictionary<string, int>();
            HashSet<int> repeatedRows = new HashSet<int>();

            for (int rowNum = 0; rowNum < checkingTable.RowCount - 1; rowNum++)
            {
                string cellValue = (string)checkingTable.Rows[rowNum].Cells[colNum].Value;
                if (cellValue == null) { continue; }
                if (existingValues.ContainsKey(cellValue))
                {
                    repeatedRows.Add(rowNum);
                    repeatedRows.Add(existingValues[cellValue]);
                }
                else
                {
                    existingValues[cellValue] = rowNum;
                }
            }

            if (repeatedRows.Count == 0)
            {
                return true;
            }

            foreach (int rowNum in repeatedRows)
            {
                checkingTable.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightPink;
            }

            return false;
        }

        private static void ResetTableColor(DataGridView dataGrid)
        {
            foreach (DataGridViewRow row in dataGrid.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell.Style.BackColor = Color.White;
                }
            }
        }

        #region Individual Cell Checks
        private bool CheckIsEmail(int rowNum, int colNum, DataGridView checkingTable, bool setWhite = false)
        {
            string value = (string)checkingTable.Rows[rowNum].Cells[colNum].Value;
            if (value == "" || value == null)
            {
                checkingTable.Rows[rowNum].Cells[colNum].Style.BackColor = Color.White;
                return true;
            }

            if (value[0] != '@')
            {
                string[] parts = value.Split('@');
                if (parts.Length != 2)
                {
                    checkingTable.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightPink;
                    return false;
                }
            }
            if (setWhite) { checkingTable.Rows[rowNum].Cells[colNum].Style.BackColor = Color.White; }
            return true;
        }

        private bool CheckEitherOr(int rowNum, int colNum1, int colNum2, DataGridView checkingTable)
        {
            string value1 = (string)checkingTable.Rows[rowNum].Cells[colNum1].Value;
            string value2 = (string)checkingTable.Rows[rowNum].Cells[colNum2].Value;
            bool passCheck = string.IsNullOrEmpty(value1) != string.IsNullOrEmpty(value2);
            if (!passCheck)
            {
                checkingTable.Rows[rowNum].Cells[colNum1].Style.BackColor = Color.LightPink;
                checkingTable.Rows[rowNum].Cells[colNum2].Style.BackColor = Color.LightPink;
            }
            return passCheck;
        }

        private bool CheckIsFolderPath(int rowNum, int colNum, DataGridView checkingTable)
        {
            string value = (string)checkingTable.Rows[rowNum].Cells[colNum].Value;
            if (value == "" || value == null)
            {
                checkingTable.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightPink;
                return false;
            }

            if (Directory.Exists(value))
            {
                checkingTable.Rows[rowNum].Cells[colNum].Style.BackColor = Color.White;
                return true;
            }
            else
            {
                checkingTable.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightPink;
                return false;
            }
        }

        private bool CheckIsNotEmpty(int rowNum, int colNum, DataGridView checkingTable)
        {
            string value = (string)checkingTable.Rows[rowNum].Cells[colNum].Value;
            if (value == "" || value == null)
            {
                checkingTable.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightPink;
                return false;
            }
            return true;
        }

        private bool CheckIsNumber(int rowNum, int colNum, DataGridView checkingTable)
        {
            string value = (string)checkingTable.Rows[rowNum].Cells[colNum].Value;
            bool isNumber = Int32.TryParse(value, out int number);
            if (!isNumber) { checkingTable.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightPink; }
            return isNumber;
        }

        private string GetGridViewName(DataGridView checkingTable)
        {
            string gridName = lastGridView.Name;
            gridName = gridName.Substring(0, gridName.Length - "GridView".Length);
            gridName = char.ToUpper(gridName[0]) + gridName.Substring(1);
            return gridName;
        }
        #endregion
        #endregion
        #endregion

        #region Basic Functions
        private void checkTableButton_Click(object sender, EventArgs e)
        {
            CheckTables();
        }
        private void deleteRow_Click(object sender, EventArgs e)
        {
            #region Warning and checks
            string gridName = GetGridViewName(lastGridView);
            if (lastGridView.SelectedRows.Count == 0) { MessageBox.Show($"No rows selected in {gridName}.", "Error"); return; }
            DialogResult res = MessageBox.Show($"Delete {lastGridView.SelectedRows.Count} rows from {gridName}?", "Confirmation", MessageBoxButtons.YesNo);
            if (res != DialogResult.Yes) { return; }
            #endregion

            foreach (DataGridViewRow row in lastGridView.SelectedRows)
            {
                lastGridView.Rows.RemoveAt(row.Index);
            }
        }

        private void clearTable_Click(object sender, EventArgs e)
        {
            string gridName = GetGridViewName(lastGridView);
            DialogResult res = MessageBox.Show($"Clear table from {gridName}?", "Confirmation", MessageBoxButtons.YesNo);
            if (res != DialogResult.Yes) { return; }
            for (int rowNum = lastGridView.RowCount - 2 ; rowNum >= 0; rowNum--)
            {
                lastGridView.Rows.RemoveAt(rowNum);
            }
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Close project without saving?",Text,MessageBoxButtons.YesNo);
            if (result != DialogResult.Yes) { return; }
            setValue = false;
            Close();
        }
        #endregion

        #region Save Value
        private void saveTable_Click(object sender, EventArgs e)
        {
            try
            {
                #region Checks
                bool checkPass = CheckTables();
                if (!checkPass)
                {
                    DialogResult res = MessageBox.Show("Error detected in data format, continue to save?", "Warning", MessageBoxButtons.YesNo);
                    if (res != DialogResult.Yes)
                    {
                        return;
                    }
                }
                #endregion

                var projectDictionary = CreateProjectDictionary();

                #region Get filePath
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "json files (*.json)|*.json";
                saveFileDialog.RestoreDirectory = true;

                DialogResult res2 = saveFileDialog.ShowDialog();
                if (res2 != DialogResult.OK) { return; }
                string filePath = saveFileDialog.FileName;
                #endregion

                WriteToJson(projectDictionary, filePath);
                MessageBox.Show($"Data saved to {filePath}", "Completed");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private Dictionary<string, object> CreateProjectDictionary()
        {
            Dictionary<string, object> projectDictionary = new Dictionary<string, object>();

            #region Sender
            HashSet<string> internalSenders = new HashSet<string>();
            foreach (DataGridViewRow row in senderGridView.Rows)
            {
                string value = (string)row.Cells[0].Value;
                if (value == null) { continue; }
                if (!internalSenders.Contains(value)) { internalSenders.Add(value.ToLower()); }
            }
            projectDictionary.Add("internalSenders", internalSenders);
            #endregion

            #region Recipient
            Dictionary<string, string> externalReferenceNames = new Dictionary<string, string>();
            for (int rowNum = 0; rowNum < recipientGridView.RowCount - 1; rowNum++)
            {
                bool isEmpty = CheckIfRowIsEmpty(recipientGridView, rowNum);
                if (isEmpty) { continue; }

                DataGridViewRow row = recipientGridView.Rows[rowNum];
                string email = ((string)row.Cells[0].Value).Trim();
                email = email.ToLower();
                string replacementText = ((string)row.Cells[1].Value).Trim();
                externalReferenceNames.Add(email, replacementText);
            }
            projectDictionary.Add("externalReferenceNames", externalReferenceNames);
            #endregion

            #region Subject
            Dictionary<string, string> subjectStrings = new Dictionary<string, string>();
            for (int rowNum = 0; rowNum < subjectGridView.RowCount - 1; rowNum++)
            {
                bool isEmpty = CheckIfRowIsEmpty(subjectGridView, rowNum);
                if (isEmpty) { continue; }

                DataGridViewRow row = subjectGridView.Rows[rowNum];
                string subjectText = ((string)row.Cells[0].Value);
                string numberOfCharacters = ((string)row.Cells[1].Value).Trim();
                subjectStrings.Add(subjectText, numberOfCharacters);
            }
            projectDictionary.Add("subjectStrings", subjectStrings);
            #endregion

            if (projectDictionary.Count == 0) { throw new ArgumentException("No value found to save"); }
            return projectDictionary;            
        }

        #endregion

        #region Load Table
        private void loadTable_Click(object sender, EventArgs e)
        {
            #region Get Json File Name
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "json files (*.json)|*.json";
            DialogResult res = dialog.ShowDialog();
            if (res == DialogResult.Cancel) { return; }
            string filePath = dialog.FileName;
            #endregion

            #region Get Project Dictionary
            Dictionary<string, object> projectDictionary = ReadJsonForHDBFilters(filePath);
            #endregion

            #region Load to table
            LoadDictionaryForHDBFilters(projectDictionary);
            #endregion
            MessageBox.Show("Data loaded", "Completed");
        }

        private Dictionary<string, object> ReadJsonForHDBFilters(string filePath)
        {
            var readDict = ReadJsonToObject<Dictionary<string, object>>(filePath);
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
                    Dictionary<string,string> dict = value.Deserialize<Dictionary<string,string>>();
                    projectDictionary.Add(key, dict);
                }
                else
                {
                    throw new ArgumentException("Unable to decompose Json value");
                }
            }

            return projectDictionary;
        }

        private void LoadDictionaryForHDBFilters(Dictionary<string, object> projectDictionary)
        {
            #region Sender
            senderGridView.Rows.Clear();

            HashSet<string> internalSenders = (HashSet<string>)projectDictionary["internalSenders"];
            foreach (string entry in internalSenders)
            {
                senderGridView.Rows.Add();
                int rowNum = senderGridView.RowCount - 2;
                senderGridView.Rows[rowNum].Cells[0].Value = entry;
            }
            senderGridView.ClearSelection();
            #endregion

            #region Recipient
            recipientGridView.Rows.Clear();

            Dictionary<string,string> externalReferenceNames = (Dictionary<string, string>)projectDictionary["externalReferenceNames"];
            foreach (KeyValuePair<string,string> entry in externalReferenceNames)
            {
                recipientGridView.Rows.Add();
                int rowNum = recipientGridView.RowCount - 2;
                recipientGridView.Rows[rowNum].Cells[0].Value = entry.Key;
                recipientGridView.Rows[rowNum].Cells[1].Value = entry.Value;
            }
            recipientGridView.ClearSelection();
            #endregion

            #region Subject
            Dictionary<string, string> subjectStrings = (Dictionary<string, string>)projectDictionary["subjectStrings"];
            foreach (KeyValuePair<string, string> entry in subjectStrings)
            {
                subjectGridView.Rows.Add();
                int rowNum = subjectGridView.RowCount - 2;
                subjectGridView.Rows[rowNum].Cells[0].Value = entry.Key;
                subjectGridView.Rows[rowNum].Cells[1].Value = entry.Value;
            }
            subjectGridView.ClearSelection();
            #endregion
        }
        #endregion

        #region Ok and Close Form
        public Dictionary<string, object> projectDictionary = null;
        public string projectName;
        public HdbExport parentForm = null;
        public bool setValue = false;
        private void okButton_Click(object sender, EventArgs e)
        {
            try
            {
                #region Checks
                // Project Names
                projectName = dispProjectName.Text;
                if (projectName == "")
                {
                    MessageBox.Show("Project name cannot be empty", "Error");
                    return;
                }

                bool contentsCheck = CheckTables();
                if (!contentsCheck)
                {
                    DialogResult result = MessageBox.Show("Error found in contents, continue to save?", "Error", MessageBoxButtons.YesNo);
                    if (result != DialogResult.Yes) { return; }
                }
                #endregion

                #region Get Project Details
                projectDictionary = CreateProjectDictionary();
                #endregion

                if (isNew && parentForm.projectTracker.ContainsKey(projectName))
                {
                    MessageBox.Show("Unable to create project as identical project name already exist in database");
                    return;
                }

                setValue = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }
        #endregion

        private void alwaysTopCheck_CheckedChanged(object sender, EventArgs e)
        {
            TopMost = alwaysTopCheck.Checked;
        }
    }
}
