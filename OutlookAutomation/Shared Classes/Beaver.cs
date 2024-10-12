using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Windows.Forms;
using System.Diagnostics; 

namespace OutlookAutomation
{
    public static class Beaver
    {
        private static readonly object _lock = new object();
        private static string filePath;
        public static bool logExist;

        public static void Initialize(string folderPath, string fileName)
        {
            filePath = Path.Combine(folderPath, fileName);
            if (Path.GetExtension(filePath) != ".txt")
            {
                throw new Exception($"Invalid filepath for output error log\nExtension type found is {Path.GetExtension(filePath)}");
            }
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            logExist = false;
        }

        //public static void InitializeForWorkbook(Microsoft.Office.Interop.Excel.Workbook workbook, string appendText = "ErrorLog")
        //{
        //    string folderPath = Path.GetDirectoryName(workbook.FullName);
        //    string fileName = Path.GetFileNameWithoutExtension(workbook.FullName) + "_" + appendText + ".txt";
        //    Initialize(folderPath, fileName);
        //}

        public static void LogError(string message)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new InvalidOperationException("Logger is not initialized. Call Logger.Initialize() with a valid file path before logging.");
            }

            try
            {
                logExist = true;
                lock (_lock)
                {
                    using (StreamWriter writer = new StreamWriter(filePath, true))
                    {
                        writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - ERROR - {message}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Logging failed: {ex.Message}");
            }
        }

        public static void LogProgress(string message)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                throw new InvalidOperationException("Logger is not initialized. Call Logger.Initialize() with a valid file path before logging.");
            }

            try
            {
                logExist = true;
                lock (_lock)
                {
                    using (StreamWriter writer = new StreamWriter(filePath, true))
                    {
                        writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Log Progress - {message}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Logging failed: {ex.Message}");
            }
        }

        public static void OpenLog()
        {
            try
            {
                Process.Start("notepad.exe", filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to open {filePath}\n\n{ex.Message}");
            }
        }

        public static void CheckLog()
        {
            if (logExist)
            {
                DialogResult result = MessageBox.Show($"Check log saved in {filePath}. \nOpen log file?", "Check log file", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    OpenLog();
                }
            }
        }
    }
}
