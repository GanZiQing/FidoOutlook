using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.IO;
using System.Windows.Forms;

namespace OutlookAutomation
{
    static class ExportUtilities
    {
        #region Json
        public static void WriteToJson(object obj, string filePath)
        {
            string contentsToWriteToFile = JsonSerializer.Serialize(obj, new JsonSerializerOptions { WriteIndented = true });
            using (StreamWriter writer = new StreamWriter(filePath, false))
            {
                writer.Write(contentsToWriteToFile);
            }
        }

        public static Dictionary<string, Dictionary<string, string>> ReadJsonToDictionary2(string filePath)
        {
            string jsonContent = File.ReadAllText(filePath);
            return JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, string>>>(jsonContent);
        }

        public static void LoadDictionaryToDataGridView(Dictionary<string, Dictionary<string, string>> projectDictionary, DataGridView dataGridView)
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

        #region For SummonForm
        //public static Dictionary<string, Dictionary<string, Dictionary<string, string>>> ReadJsonToProjectTracker(string filePath)
        //{
        //    // Read the JSON content from the file
        //    string jsonContent = File.ReadAllText(filePath);

        //    // Deserialize the JSON content into the specified nested dictionary format
        //    return JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, Dictionary<string, string>>>>(jsonContent);
        //}

        public static Dictionary<TKey, TValue> ReadJsonToDictionary<TKey, TValue>(string filePath)
        {
            // Read the JSON content from the file
            string jsonContent = File.ReadAllText(filePath);

            // Deserialize the JSON content into the specified dictionary format
            return JsonSerializer.Deserialize<Dictionary<TKey, TValue>>(jsonContent);
        }

        public static T ReadJsonToObject<T>(string filePath)
        {
            // Read the JSON content from the file
            string json = File.ReadAllText(filePath);

            // Deserialize the JSON into the specified type
            return JsonSerializer.Deserialize<T>(json);
        }

        

        #endregion
    }
}
