using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;
using ConvertLibrary.Models;
using static ConvertLibrary.BusinessLogic.DelimiterManager;
using System.Linq;

namespace ConvertLibrary.BusinessLogic
{
    public class FileConverter
    {
        // Print
        public static void PrintTableData(Table table)
        {
            foreach (var row in table.Rows)
            {
                string rowData = string.Join(", ", row);
                Console.WriteLine(rowData);
            }
        }

        // Read
        public Table ReadExcelFile(string xlsFilePath, int skipRange, CsvDelimiter selectedDelimiter)
        {
            var table = new Table(new List<string[]>(), new List<string[]>(), string.Empty, 0, skipRange, selectedDelimiter);

            using (var stream = File.Open(xlsFilePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                int rowCount = 0;
                while (reader.Read())
                {
                    if (rowCount < skipRange)
                    {
                        rowCount++;
                        continue; // Skip rows based on skipRange
                    }

                    var row = new List<string>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        row.Add(reader.GetValue(i)?.ToString());
                    }
                    table.Rows.Add(row.ToArray());
                }
            }
            return table;
        }

        // Write
        public void WriteTableToCsv(Table table, string csvFilePath, CsvDelimiter delimiter)
        {
            char delimiterChar = DelimiterManager.GetDelimiterChar(delimiter);
            string delimiterString = delimiterChar.ToString(); // Umwandlung von char in string

            using (var writer = new StreamWriter(csvFilePath))
            {
                foreach (var row in table.Rows)
                {
                    string csvRow = string.Join(delimiterString, row);
                    writer.WriteLine(csvRow);
                }
            }
        }

        //Convert
        public void ConvertXlsToCsv(string xlsFilePath, int skipRange, string sheetName, CsvDelimiter selectedDelimiter)
        {
            string csvFolderName = "Results";
            string solutionDirectory = Directory.GetParent(Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName).FullName;
            string csvFolderPath = Path.Combine(solutionDirectory, "XlsXMht2Csv", "bin", "Debug", csvFolderName);
            Directory.CreateDirectory(csvFolderPath);

            string xlsFileName = Path.GetFileName(xlsFilePath);
            
            //string csvFileName = Path.ChangeExtension(xlsFileName, ".csv");
            string csvFileName = GenerateNewCsvFilePath(xlsFileName); // Time Stamp Methode GenerateNewCsvFilePath

            string csvFilePath = Path.Combine(csvFolderPath, csvFileName);

            CsvDelimiter delimiter = selectedDelimiter;

            Table table = ReadExcelFile(xlsFilePath, skipRange, selectedDelimiter);

            // Print the table data to console for verification
            PrintTableData(table);

            try
            {
                // Write the table data to the CSV file
                WriteTableToCsv(table, csvFilePath, delimiter);

                Console.WriteLine("Die Konvertierung wurde abgeschlossen.");
                Console.WriteLine($"Die CSV-Datei wurde unter folgendem Pfad gespeichert: {csvFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fehler beim Schreiben der CSV-Datei: {ex.Message}");
            }
        }
        
        // Time Stamp
        private string GenerateNewCsvFilePath(string originalPath)
        {
            string directory = Path.GetDirectoryName(originalPath);
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalPath);
            string extension = Path.GetExtension(originalPath);
            string newFileName = $"{fileNameWithoutExtension}_{DateTime.Now:yyyyMMddHHmm}{extension}";
            return Path.Combine(directory, newFileName);
        }

    }
}