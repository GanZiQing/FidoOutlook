using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;

namespace OutlookAutomation
{
    public partial class CustomRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        #region User Control Manager
        private UserControl CreateUserControl(string taskPaneName)
        {
            UserControl userControl;

            switch (taskPaneName)
            {
                case "PrintPane":
                    {
                        userControl = (new PrintPane());
                        break;
                    }
                default:
                    {
                        throw new Exception($"Pane type {taskPaneName} not found");
                    }
            }
            return userControl;
        }
        #endregion

        #region Launch Buttons 
        private void printPane_Click(object sender, RibbonControlEventArgs e)
        {
            TogglePane("PrintPane");
        }
        #endregion

        #region Task Pane Manager

        private void TogglePane(string taskPaneName)
        {
            CustomTaskPane thisPane = AddOrGetSingleTaskPane(taskPaneName);
            thisPane.Visible = !thisPane.Visible;
        }

        Dictionary<string, CustomTaskPane> paneTypeDictionary = new Dictionary<string, CustomTaskPane>();
        private CustomTaskPane AddOrGetSingleTaskPane(string taskPaneName)
        {
            // Get Task Panes if Exist, else create new one
            if (!paneTypeDictionary.ContainsKey(taskPaneName))
            {
                UserControl userControl = CreateUserControl(taskPaneName);
                int controlWidth = userControl.Width + 10;
                CustomTaskPane paneValue = Globals.ThisAddIn.CustomTaskPanes.Add(userControl, taskPaneName);
                paneValue.Width = controlWidth;
                paneTypeDictionary[taskPaneName] = paneValue;

            }
            return paneTypeDictionary[taskPaneName];
        }
        #endregion
    }
}

