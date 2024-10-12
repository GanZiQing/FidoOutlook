using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using Action = System.Action;
using Exception = System.Exception;

namespace OutlookAutomation
{
    static class GlobalUtilities
    {
        #region FileNames and Paths
        public static string GetAvailableFileName(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return filePath;
            }

            // Get the directory, filename, and extension
            string directory = Path.GetDirectoryName(filePath);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);

            int fileNumber = 1;

            // Continue looping until a file with the new name does not exist
            string newFilePath = filePath;
            while (File.Exists(newFilePath))
            {
                newFilePath = Path.Combine(directory, $"{fileNameWithoutExtension} ({fileNumber}){extension}");
                fileNumber++;
                if (fileNumber >= 100)
                {
                    throw new Exception($"File already exist and unable to find new file name.\n{filePath}");
                }
            }

            return newFilePath;
        }

        public static string GetAvailableFolderName(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                return folderPath;
            }

            // Get the directory, filename, and extension
            string parentDirectory = Path.GetDirectoryName(folderPath);
            string currentFolder = Path.GetFileName(folderPath);
            int folderNumber = 1;

            // Continue looping until a file with the new name does not exist
            string newFolderPath = folderPath;
            while (Directory.Exists(newFolderPath))
            {
                newFolderPath = Path.Combine(parentDirectory,currentFolder+$" ({folderNumber})");
                folderNumber++;
                if (folderNumber >= 100)
                {
                    throw new Exception($"Folder already exist and unable to find new file name.\n{folderPath}");
                }
            }

            return newFolderPath;
        }

        public static string ShortenFolderPath(string folderPath, int maxFolderLength = 256)
        {
            string parentDirectory = Path.GetDirectoryName(folderPath);
            string currentFolder = Path.GetFileName(folderPath);

            string newFolder = Substring2(currentFolder, maxFolderLength);
            string newPath = Path.Combine(parentDirectory, newFolder);

            if (newPath.Length > 256)
            {
                int remainingCharacters = 256 - parentDirectory.Length - 1;
                if (remainingCharacters <= 0)
                {
                    throw new Exception($"File Path exceeds maximum characters");
                }
                newFolder = Substring2(newFolder, remainingCharacters);
                newPath = Path.Combine(parentDirectory, newFolder);
            }
            return newPath;
        }

        public static string Substring2(string text, int maxLength)
        {
            if (text.Length < maxLength)
            {
                return text;
            }
            else
            {
                return text.Substring(0, maxLength);
            }
        }

        public static string SanitiseFileName(string inputFileName)
        {
            HashSet<char> invalidChars = new HashSet<char>(Path.GetInvalidFileNameChars());

            StringBuilder sanitizedFileName = new StringBuilder();
            foreach (char c in inputFileName)
            {
                if (!invalidChars.Contains(c))
                {
                    sanitizedFileName.Append(c);
                }
            }

            return sanitizedFileName.ToString();
        }

        public static string SanitiseFilePath(string inputFilePath)
        {
            HashSet<char> invalidChars = new HashSet<char>(Path.GetInvalidPathChars());

            StringBuilder sanitizedFilePath = new StringBuilder();
            foreach (char c in inputFilePath)
            {
                if (!invalidChars.Contains(c))
                {
                    sanitizedFilePath.Append(c);
                }
            }

            return sanitizedFilePath.ToString();
        }
        #endregion



        public static int ConvertToProgress(int currentProgress, int maxProgress)
        {
            if (maxProgress == 0) { return 0; }
            double progressDouble = Convert.ToDouble(currentProgress) / Convert.ToDouble(maxProgress) * 100;
            int progress = Convert.ToInt32(progressDouble);
            if (progress > 100) { progress = 100; }
            return progress;
        }

        #region Word Related
        public static Word.Application GetWordApp()
        {
            Word.Application wordApp = null;
            try
            {
                // Try to get the active Word application
                wordApp = (Word.Application)Marshal.GetActiveObject("Word.Application");

                // Warn user if there is an active application
                DialogResult dialogResult = MessageBox.Show("There is an active Word instance. Please save your open documents. Continue?", "Warning", MessageBoxButtons.YesNo);
                if (dialogResult != DialogResult.Yes)
                {
                    throw new Exception("Terminated by user");
                }
            }
            catch (COMException)
            {
                // No active Word application found, create a new instance
                wordApp = new Word.Application();
            }

            return wordApp;
        }
        #endregion

        #region Retry
        public static void RetryIfBusy(Action action, int maxRetries = 10, int delayMs = 500)
        {
            int retryCount = 0;
            while (retryCount < maxRetries)
            {
                try
                {
                    action();
                    return; // Success, exit the loop
                }
                catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x8001010A)) // RPC_E_SERVERCALL_RETRYLATER
                {
                    retryCount++;
                    if (retryCount >= maxRetries)
                        throw; // Re-throw after max retries
                    System.Threading.Thread.Sleep(delayMs); // Wait before retrying
                }
            }
        }

        public static void RetryForInspector(Action action, Action finallyAction,int maxRetries = 10, int delayMs = 500)
        {
            int retryCount = 0;
            while (retryCount < maxRetries)
            {
                try
                {
                    action();
                    return; // Success, exit the loop
                }
                catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x8001010A)) // RPC_E_SERVERCALL_RETRYLATER
                {
                    retryCount++;
                    if (retryCount >= maxRetries)
                        throw; // Re-throw after max retries
                    System.Threading.Thread.Sleep(delayMs); // Wait before retrying
                }
                finally
                {

                }
            }
        }
        #endregion

        public static void DeleteFolderIfEmpty(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                // Check if the folder has any files or subdirectories
                if (Directory.GetFiles(folderPath).Length == 0 && Directory.GetDirectories(folderPath).Length == 0)
                {
                    Directory.Delete(folderPath);
                }
            }
        }
    }

    static class OutlookUtilities
    {
        public static Folder GetFolderFromFolder(Folder baseFolder, string name, bool createIfNull = false)
        {
            Folder subFolder = null;
            try
            {
                subFolder = baseFolder.Folders[name] as Folder;
            }
            catch (COMException ex) when (ex.ErrorCode == - 2147221233)
            {
                if (!createIfNull) { return null; }
                baseFolder.Folders.Add(name);
            }

            return baseFolder.Folders[name] as Folder;
        }
    }

    class CustomFolderBrowser
    {
        OpenFileDialog dialog = new OpenFileDialog();
        public CustomFolderBrowser()
        {
            dialog.ValidateNames = false;  // Allows selecting folders
            dialog.Filter = "Folders|*. ";
            dialog.CheckFileExists = false;
            dialog.CheckPathExists = true;
            dialog.FileName = "Select Folder";  // Fake name to allow folder selection
        }

        public string folderPath = null;
        public DialogResult ShowDialog()
        {
            DialogResult dialogResult = dialog.ShowDialog();
            if (dialogResult == DialogResult.OK)
            {
                string test = dialog.FileName;
                folderPath = Path.GetDirectoryName(dialog.FileName);
            }
            return dialogResult;
        }

        public string GetFolderPath()
        {
            if (folderPath == null)
            {
                throw new Exception("Folder path is not set");
            }
            return folderPath;
        }
    }

}
