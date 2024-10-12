using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace OutlookAutomation
{
    //class CustomFolderBrowser
    //{
    //    #region Sample Usage
    //    //if (folderPath == "")
    //    //{
    //    //    CustomFolderBrowser customFolderBrowser = new CustomFolderBrowser();
    //    //    DialogResult res = customFolderBrowser.ShowDialog();
    //    //    if (res == DialogResult.Cancel) { return; }
    //    //    folderPath = customFolderBrowser.folderPath;
    //    //}
    //    #endregion

    //    OpenFileDialog dialog = new OpenFileDialog();
    //    public CustomFolderBrowser()
    //    {
    //        dialog.ValidateNames = false;  // Allows selecting folders
    //        dialog.Filter = "Folders|*. ";
    //        dialog.CheckFileExists = false;
    //        dialog.CheckPathExists = true;
    //        dialog.FileName = "Select Folder";  // Fake name to allow folder selection
    //    }

    //    public string folderPath = null;
    //    public DialogResult ShowDialog()
    //    {
    //        DialogResult dialogResult = dialog.ShowDialog();
    //        if (dialogResult == DialogResult.OK)
    //        {
    //            string test = dialog.FileName;
    //            folderPath = Path.GetDirectoryName(dialog.FileName);
    //        }
    //        return dialogResult;
    //    }

    //    public string GetFolderPath()
    //    {
    //        if (folderPath == null)
    //        {
    //            throw new Exception("Folder path is not set");
    //        }
    //        return folderPath;
    //    }
    //}
}
