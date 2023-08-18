using ConvertLibrary.Models;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ConvertLibrary.BusinessLogic.DelimiterManager;

namespace ConvertLibrary.BusinessLogic
{
    public class TableHelper
    {
        public static int GetSheetCount(string xlsFilePath)
        {
            int sheetCount = 0;

            using (var stream = File.Open(xlsFilePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                while (reader.NextResult())
                {
                    sheetCount++;
                }
            }

            return sheetCount;
        }

        public static List<Table> ReadAllSheets(string xlsFilePath, CsvDelimiter delimiter, int skipRange)
        {
            List<Table> tables = new List<Table>();
            int sheetCount = GetSheetCount(xlsFilePath);

            for (int sheetSet = 0; sheetSet < sheetCount; sheetSet++)
            {
                Table table = ReadSheet(xlsFilePath, sheetSet, delimiter, skipRange);
                tables.Add(table);
            }

            return tables;
        }

        public static string GetSheetName(string xlsFilePath, int sheetSet)
        {
            using (var stream = File.Open(xlsFilePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                while (reader.Read())
                {
                    int sheetIndex = reader.Depth;
                    if (sheetIndex == sheetSet)
                    {
                        return reader.Name;
                    }
                }
            }
            return null; // Sheet not found
        }

        public static Table ReadExcelSheet(string xlsFilePath, int sheetSet, int skipRange, CsvDelimiter delimiter)
        {
            var rows = ReadSheetRows(xlsFilePath, sheetSet, skipRange);
            string sheetName = GetSheetName(xlsFilePath, sheetSet); // Get the sheet name
            int totalSheetCount = GetSheetCount(xlsFilePath); // Get the total sheet count
            return new Table(rows, null, sheetName, totalSheetCount, skipRange, delimiter);
        }

        public static Table ReadSheet(string xlsFilePath, int sheetSet, CsvDelimiter delimiter, int skipRange)
        {
            string sheetName = GetSheetName(xlsFilePath, sheetSet);
            List<string[]> rows = ReadSheetRows(xlsFilePath, sheetSet, skipRange);
            return new Table(rows, null, sheetName, sheetSet, skipRange, delimiter);
        }

        private static List<string[]> ReadSheetRows(string xlsFilePath, int sheetSet, int skipRange)
        {
            var rows = new List<string[]>();

            using (var stream = File.Open(xlsFilePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                int currentSheet = -1;
                int rowCount = 0;

                while (reader.Read())
                {
                    if (reader.Depth != currentSheet)
                    {
                        currentSheet = reader.Depth;
                        rowCount = 0;
                    }

                    if (currentSheet == sheetSet && rowCount >= skipRange)
                    {
                        var row = new List<string>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            row.Add(reader.GetValue(i)?.ToString());
                        }
                        rows.Add(row.ToArray());
                    }

                    rowCount++;
                }
            }

            return rows;
        }

        public static void WriteAllTablesToCsv(List<Table> tables)
        {
            foreach (Table table in tables)
            {
                WriteTableToCsv(table);
            }
        }

        public static void WriteTableToCsv(Table table)
        {
            string csvFilePath = GenerateCsvFilePath(table.SheetName);
            using (StreamWriter writer = new StreamWriter(csvFilePath))
            {
                foreach (string[] row in table.Rows)
                {
                    string csvRow = string.Join(table.Delimiter.ToString(), row);
                    writer.WriteLine(csvRow);
                }
            }
        }

        private static string GenerateCsvFilePath(string sheetName)
        {
            string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results");
            Directory.CreateDirectory(directory);
            string csvFileName = $"{sheetName}_{DateTime.Now:yyyyMMddHHmm}.csv";
            return Path.Combine(directory, csvFileName);
        }


    }
}
