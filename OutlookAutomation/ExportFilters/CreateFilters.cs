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

namespace OutlookAutomation
{
    public partial class CreateFilters : Form
    {
        #region Init
        public bool isNew = false;
        public CreateFilters(AdvanceExport parentForm) 
        {
            InitializeComponent();
            CancelButton = cancelButton;
            StartPosition = FormStartPosition.CenterScreen;
            this.parentForm = parentForm;
            isNew = true;
            Text = "Add New Project";
        }

        public string originalProjectName;
        public CreateFilters(AdvanceExport parentForm, string projectName, Dictionary<string,Dictionary<string,string>> tableData)
        {
            InitializeComponent();
            dispProjectName.Text = projectName;
            originalProjectName = projectName;
            Text = originalProjectName;
            LoadDictionaryToDataGridView(tableData, dataGridView);
        }

        #region Events
        private void SubscribeToEvents()
        {
            //dataGrid.CellValidating += new DataGridViewCellValidatingEventHandler(CellLeave);
            //dataGrid.CellLeave += new DataGridViewCellEventHandler(CellLeaveEvent);
        }
        private void CellValidatingEvent(object sender, DataGridViewCellValidatingEventArgs e)
        {
            ValidateOneCell(e.RowIndex, e.ColumnIndex);
        }
        private void CellLeaveEvent(object sender, DataGridViewCellEventArgs e)
        {
            ValidateOneCell(e.RowIndex, e.ColumnIndex);
        }
        #endregion
        #endregion

        #region Data Validation


        private bool CheckTable()
        {
            bool failCheck = false;

            #region Check Empty
            bool isEmpty = CheckIfTableIsEmpty(dataGridView);
            if (isEmpty)
            {
                failCheck = true;
                MessageBox.Show("Table is empty", "Error");
                return (!failCheck);
            }
            #endregion

            ResetTableColor(dataGridView);

            #region Check Names
            for (int i = 0; i < dataGridView.RowCount - 1; i++)
            {
                bool[] checks = new bool[4];
                // Check name
                checks[0] = CheckIsNotEmpty(i, 0);
                // Check email
                checks[1] = CheckIsEmail(i, 1);
                checks[2] = CheckIsEmail(i, 2);

                // Check folder
                checks[3] = CheckIsFolderPath(i, 3);

                foreach (bool check in checks)
                {
                    if (!check) { failCheck = true; }
                }
            }
            #endregion
            HighlightEmptyRows(dataGridView);
            HighlightRepeatedRows(dataGridView,0);
            return !failCheck;
        }

        private static bool CheckIfTableIsEmpty(DataGridView dataGrid)
        {
            for (int rowNum = 0; rowNum < dataGrid.RowCount - 1; rowNum++)
            {
                if (!CheckIfRowIsEmpty(dataGrid, rowNum))
                {
                    return false;
                }
            }
            return true;
        }

        private static bool CheckIfRowIsEmpty(DataGridView dataGrid, int rowNum)
        {
            for (int colNum = 0; colNum < dataGrid.ColumnCount; colNum++)
            {
                string value = (string)dataGrid.Rows[rowNum].Cells[colNum].Value;
                if (!(value == "" || value == null))
                {
                    return false;
                }
            }
            return true;
        }

        private static void HighlightEmptyRows(DataGridView dataGrid)
        {
            List<int> emptyRowNums = new List<int>();
            for (int rowNum = 0; rowNum < dataGrid.RowCount - 1; rowNum++)
            {
                if (CheckIfRowIsEmpty(dataGrid, rowNum))
                {
                    emptyRowNums.Add(rowNum);
                }
            }

            if (emptyRowNums.Count == 0) { return; }

            DialogResult res = MessageBox.Show("Empty rows found, delete rows?", "Empty Rows", MessageBoxButtons.YesNo);
            emptyRowNums.Reverse();
            if (res == DialogResult.Yes)
            {
                foreach (int rowNum in emptyRowNums)
                {
                    dataGrid.Rows.RemoveAt(rowNum);
                }
            }
            else
            {
                foreach (int rowNum in emptyRowNums)
                {
                    for (int colNum = 0; colNum < dataGrid.ColumnCount; colNum++)
                    {
                        dataGrid.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightGray;
                    }
                }
            }
        }
        
        private static bool HighlightRepeatedRows(DataGridView dataGrid, int colNum)
        {
            Dictionary<string, int> existingValues = new Dictionary<string, int>();
            HashSet<int> repeatedRows = new HashSet<int>();

            for (int rowNum = 0; rowNum < dataGrid.RowCount - 1; rowNum++)
            {
                string cellValue = (string)dataGrid.Rows[rowNum].Cells[colNum].Value;
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
                dataGrid.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightPink;
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
        private void ValidateOneCell(int rowNum, int colNum)
        {
            if (colNum == 1 || colNum == 2)
            {
                CheckIsEmail(rowNum, colNum);
            }
            else if (colNum == 3)
            {
                CheckIsFolderPath(rowNum, colNum);
            }
        }

        private bool CheckIsEmail(int rowNum, int colNum)
        {
            string value = (string)dataGridView.Rows[rowNum].Cells[colNum].Value;
            if (value == "" || value == null)
            {
                dataGridView.Rows[rowNum].Cells[colNum].Style.BackColor = Color.White;
                return true;
            }

            if (value[0] != '@')
            {
                string[] parts = value.Split('@');
                if (parts.Length != 2)
                {
                    dataGridView.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightPink;
                    return false;
                }
            }
            dataGridView.Rows[rowNum].Cells[colNum].Style.BackColor = Color.White;
            return true;
        }

        private bool CheckIsFolderPath(int rowNum, int colNum)
        {
            string value = (string)dataGridView.Rows[rowNum].Cells[colNum].Value;
            if (value == "" || value == null)
            {
                dataGridView.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightPink;
                return false;
            }

            if (Directory.Exists(value))
            {
                dataGridView.Rows[rowNum].Cells[colNum].Style.BackColor = Color.White;
                return true;
            }
            else
            {
                dataGridView.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightPink;
                return false;
            }
        }

        private bool CheckIsNotEmpty(int rowNum, int colNum)
        {
            string value = (string)dataGridView.Rows[rowNum].Cells[colNum].Value;
            if (value == "" || value == null)
            {
                dataGridView.Rows[rowNum].Cells[colNum].Style.BackColor = Color.LightPink;
                return false;
            }
            return true;
        }
        #endregion
        #endregion

        #region Basic Functions
        private void checkTableButton_Click(object sender, EventArgs e)
        {
            CheckTable();
        }

        private void deleteRow_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView.SelectedRows)
            {
                dataGridView.Rows.RemoveAt(row.Index);
            }
        }

        private void clearTable_Click(object sender, EventArgs e)
        {
            for (int rowNum = dataGridView.RowCount - 2 ; rowNum >= 0; rowNum--)
            {
                dataGridView.Rows.RemoveAt(rowNum);
            }
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Close project without saving?",Text,MessageBoxButtons.YesNo);
            if (result != DialogResult.Yes) { return; }
            Close();
        }
        #endregion

        #region Save Value
        private void saveTable_Click(object sender, EventArgs e)
        {
            try
            {
                #region Checks
                bool checkPass = CheckTable();
                if (!checkPass)
                {
                    DialogResult res = MessageBox.Show("Error detected in data format, continue to save?", "Warning", MessageBoxButtons.YesNo);
                    if (res != DialogResult.Yes)
                    {
                        return;
                    }
                }
                #endregion


                #region Get folderPath
                CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
                DialogResult res2 = customFolderBrowser.ShowDialog();
                if (res2 == DialogResult.Cancel) { return; }
                string folderPath = customFolderBrowser.folderPath;
                folderPath = Path.Combine(folderPath, "TestProject.json"); // Comeback to update this to use save file dilogue
                #endregion

                var projectDictionary = CreateProjectDictionary(dataGridView);
                WriteProjectToJson(projectDictionary, folderPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private Dictionary<string, Dictionary<string, string>> CreateProjectDictionary(DataGridView dataGridView)
        {
            Dictionary<string, Dictionary<string, string>> projectDictionary = new Dictionary<string, Dictionary<string, string>>();

            for (int rowNum = 0; rowNum < dataGridView.RowCount - 1; rowNum++)
            {
                bool isEmpty = CheckIfRowIsEmpty(dataGridView, rowNum);
                if (isEmpty) { continue; }
                //SingleFilter singleFilter = new SingleFilter(dataGridView, i);

                #region Create Dictionary for Each Row
                Dictionary<string, string> contents = new Dictionary<string, string>();
                DataGridViewRow row = dataGridView.Rows[rowNum];

                for (int colNum = 0; colNum < dataGridView.ColumnCount; colNum++)
                {
                    string colName = dataGridView.Columns[colNum].Name;
                    string cellValue = (string)row.Cells[colName].Value;

                    if (cellValue == null) { cellValue = ""; }

                    if (colName == "sender" || colName == "recipient")
                    {
                        cellValue = cellValue.ToLower();
                    }
                    contents[colName] = cellValue;
                }
                
                #endregion

                projectDictionary[contents["name"]] = contents;
            }

            if (projectDictionary.Count == 0) { throw new ArgumentException("No value found to save"); }
            return projectDictionary;            
        }

        private void WriteProjectToJson(Dictionary<string, Dictionary<string, string>> projectDictionary, string filePath)
        {
            string projectName = dispProjectName.Text;
            if (projectName == "") { throw new ArgumentException("Project name cannot be empty"); }

            WriteToJson(projectDictionary, filePath);
            MessageBox.Show($"Data saved to {filePath}", "Completed");
        }


        #endregion

        #region Json

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
            Dictionary<string, Dictionary<string, string>> projectDictionary = ReadJsonToDictionary2(filePath);
            #endregion

            #region Load to table
            LoadDictionaryToDataGridView(projectDictionary, dataGridView);
            #endregion
            MessageBox.Show("Data loaded", "Completed");
        }

        public static void LoadSummonDictionary(Dictionary<string, Dictionary<string, string>> projectDictionary, DataGridView dataGridView)
        {
            dataGridView.Rows.Clear();

            foreach (Dictionary<string, string> rowEntry in projectDictionary.Values)
            {
                dataGridView.Rows.Add();
                int rowNum = dataGridView.RowCount - 2;
                foreach (KeyValuePair<string, string> colEntry in rowEntry)
                {
                    string colName = colEntry.Key;
                    string colValue = colEntry.Value;

                    dataGridView.Rows[rowNum].Cells[colName].Value = colValue;
                }
            }
        }

        #endregion

        #region Close

        #endregion

        public Dictionary<string, Dictionary<string, string>> projectDictionary = null;
        public string projectName;
        public AdvanceExport parentForm = null;
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

                bool contentsCheck = CheckTable();
                if (!contentsCheck)
                {
                    DialogResult result = MessageBox.Show("Error found in contents, continue to save?", "Error", MessageBoxButtons.YesNo);
                    if (result != DialogResult.Yes) { return; }
                    
                }
                #endregion
                #region Get Project Details

                projectDictionary = CreateProjectDictionary(dataGridView);
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

        private void alwaysTopCheck_CheckedChanged(object sender, EventArgs e)
        {
            TopMost = alwaysTopCheck.Checked;
        }
    }
}
