using ConvertLibrary.BusinessLogic;
using ConvertLibrary.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using static ConvertLibrary.BusinessLogic.DelimiterManager;

namespace XlsXMht2Csv
{
    public class Program
    {
        static void Main(string[] args)
        {
            // input // XlsXMht2Csv/bin/Debug/Data
            string xlsFileName = "TEST_XLS.xls";
            string testDataFolder = "Data";
            string xlsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, testDataFolder, xlsFileName);

            Console.WriteLine("Generierter Pfad zur XLS-Datei: " + xlsFilePath);
            Console.WriteLine("Existiert der Pfad? " + File.Exists(xlsFilePath));

            // output // XlsXMht2Csv/bin/Debug/Results
            string csvFolderName = "Results";
            string solutionDirectory = Directory.GetParent(Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName).FullName;
            string csvFolderPath = Path.Combine(solutionDirectory, "XlsXMht2Csv", "bin", "Debug", csvFolderName);
            string xlsFileNameWithoutExtension = Path.GetFileNameWithoutExtension(xlsFileName);
            string csvFileName = $"{xlsFileNameWithoutExtension}.csv";
            string csvFilePath = Path.Combine(csvFolderPath, csvFileName);

            Console.WriteLine("Generierter Pfad zur CSV-Datei: " + csvFilePath);
            Console.WriteLine("Existiert der Pfad? " + Directory.Exists(csvFolderPath));
            Console.WriteLine("Existiert die CSV-Datei? " + File.Exists(csvFilePath));


            ConvertXls();
            //CheckPaths();
            //CheckPathsAndDebug();
            //ChooseSeparator();
            //TestCsvPath();
            //TestSingleSheetXls();
            //TestMultiSheetXls();
            //PrintTableData(Table table)
        }
        public static void CheckPaths()
        {
            string xlsFileName = "TEST_XLS.xls";
            string testDataFolder = "Data";
            string xlsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, testDataFolder, xlsFileName);

            Console.WriteLine("Generierter Pfad zur XLS-Datei: " + xlsFilePath);
            Console.WriteLine("Existiert der Pfad? " + File.Exists(xlsFilePath));

            if (!File.Exists(xlsFilePath))
            {
                Console.WriteLine("Die XLS-Datei existiert nicht.");
                Console.WriteLine("Drücke eine beliebige Taste zum Beenden...");
                Console.ReadKey();
                Environment.Exit(1);
            }

            string csvFolderName = "Results";
            string solutionDirectory = Directory.GetParent(Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName).FullName;
            string csvFolderPath = Path.Combine(solutionDirectory, "XlsXMht2Csv", "bin", "Debug", csvFolderName);
            string xlsFileNameWithoutExtension = Path.GetFileNameWithoutExtension(xlsFileName);
            string csvFileName = $"{xlsFileNameWithoutExtension}.csv";
            string csvFilePath = Path.Combine(csvFolderPath, csvFileName);

            Console.WriteLine("Generierter Pfad zur CSV-Datei: " + csvFilePath);
            Console.WriteLine("Existiert der Pfad? " + Directory.Exists(csvFolderPath));
            Console.WriteLine("Existiert die CSV-Datei? " + File.Exists(csvFilePath));

            Console.ReadKey();
        }

        public static void TestCsvPath()
        {
            string csvFilePath = @"C:\Users\georg.schweizer\source\repos\REPO\XlsXMht2Csv_Plugin\XlsXMht2Csv\bin\Debug\Results\TEST_XLS.csv";

            if (!File.Exists(csvFilePath))
            {
                Console.WriteLine($"Die CSV-Datei {csvFilePath} existiert nicht.");
                Console.WriteLine("Drücke eine beliebige Taste zum Beenden...");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("Die CSV-Datei wurde gefunden.");
            // Weitere Schritte zur Verarbeitung der CSV-Datei können hier hinzugefügt werden
            Console.ReadLine();
        }

        public static void ConvertXls()
        {
            string xlsFileName = "TEST_XLS.xls";
            string testDataFolder = "Data";
            string xlsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, testDataFolder, xlsFileName);

            Console.WriteLine("Generierter Pfad zur XLS-Datei: " + xlsFilePath);
            Console.WriteLine("Existiert der Pfad? " + File.Exists(xlsFilePath));

            if (!File.Exists(xlsFilePath))
            {
                Console.WriteLine("Die XLS-Datei existiert nicht.");
                Console.WriteLine("Drücke eine beliebige Taste zum Beenden...");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("Die XLS-Datei wurde gefunden.");

            FileConverter converter = new FileConverter();

            Console.WriteLine("Gib die Anzahl der zu überspringenden Zeilen ein:");
            int skipRange = int.Parse(Console.ReadLine());

            Console.WriteLine("Gib den Namen des Arbeitsblatts ein:");
            string sheetName = Console.ReadLine();

            Console.WriteLine("Gib das Trennzeichen (0 für Komma, 1 für Semikolon) ein:");
            CsvDelimiter selectedDelimiter = (CsvDelimiter)int.Parse(Console.ReadLine());
            Console.WriteLine($"Zeichentrennung: {selectedDelimiter.ToString()} ");
            string csvFolderName = "Results";
            string solutionDirectory = Directory.GetParent(Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName).FullName;
            string csvFolderPath = Path.Combine(solutionDirectory, "XlsXMht2Csv", "bin", "Debug", csvFolderName);
            string xlsFileNameWithoutExtension = Path.GetFileNameWithoutExtension(xlsFileName);
            string csvFileName = $"{xlsFileNameWithoutExtension}.csv";
            string csvFilePath = Path.Combine(csvFolderPath, csvFileName);
            Directory.CreateDirectory(csvFolderPath);

            converter.ConvertXlsToCsv(xlsFilePath, skipRange, sheetName, selectedDelimiter);
            Console.WriteLine("Konvertiere XLS-Datei in CSV...");

            Console.WriteLine("Convertierung abgeschlossen");
            Console.WriteLine($"Filepath: {csvFilePath}");
            Console.WriteLine($"Filepath: {csvFolderPath}");

            Console.ReadLine();
        }

        public static void ChooseSeparator()
        {
            string xlsFileName = "TEST_XLS.xls";
            string testDataFolder = "Data";
            string xlsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, testDataFolder, xlsFileName);

            Console.WriteLine("Generierter Pfad zur XLS-Datei: " + xlsFilePath);
            Console.WriteLine("Existiert der Pfad? " + File.Exists(xlsFilePath));

            if (!File.Exists(xlsFilePath))
            {
                Console.WriteLine("Die XLS-Datei existiert nicht.");
                Console.WriteLine("Drücke eine beliebige Taste zum Beenden...");
                Console.ReadKey();
                return;
            }

            Console.WriteLine("Die XLS-Datei wurde gefunden.");

            FileConverter converter = new FileConverter();

            Console.WriteLine("Gib die Anzahl der zu überspringenden Zeilen ein:");
            int skipRange = int.Parse(Console.ReadLine());

            Console.WriteLine("Gib den Namen des Arbeitsblatts ein:");
            string sheetName = Console.ReadLine();

            Console.WriteLine("Wähle das Trennzeichen (0 für Komma, 1 für Semikolon):");
            CsvDelimiter selectedDelimiter = (CsvDelimiter)int.Parse(Console.ReadLine());

            Console.WriteLine("Konvertiere XLS-Datei in CSV...");

            Console.WriteLine("Konvertierung abgeschlossen.");

            string csvFolderName = "Results";
            string solutionDirectory = Directory.GetParent(Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName).FullName;
            string csvFolderPath = Path.Combine(solutionDirectory, "XlsXMht2Csv", "bin", "Debug", csvFolderName);
            string xlsFileNameWithoutExtension = Path.GetFileNameWithoutExtension(xlsFileName);
            string csvFileName = $"{xlsFileNameWithoutExtension}.csv";
            string csvFilePath = Path.Combine(csvFolderPath, csvFileName);

            FileConverter fc = new FileConverter();
            //string csvFilePath = Path.Combine(csvFolderPath, csvFileName);
            //fc.ReadCsvFileAndDisplay(csvFilePath);

            Console.ReadLine();
        }

        public static void TestSingleSheetXls()
        {
            string xlsFilePath = "Path_To_Your_XLS_File.xls";
            int skipRange = 1; // Number of rows to skip

            FileConverter converter = new FileConverter();
            Table table = converter.ReadExcelFile(xlsFilePath, skipRange, CsvDelimiter.Comma); // Change delimiter if needed
            PrintTableData(table);
        }

        public static void TestMultiSheetXls()
        {
            string xlsFilePath = "Path_To_Your_XLS_File.xls";

            int sheetCount = TableHelper.GetSheetCount(xlsFilePath);

            for (int sheetNo = 0; sheetNo < sheetCount; sheetNo++)
            {
                Table table = TableHelper.ReadSheet(xlsFilePath, sheetNo, CsvDelimiter.Comma, 0); // Change delimiter and skipRange if needed
                PrintTableData(table);
            }
        }

        public static void PrintTableData(Table table)
        {
            foreach (var row in table.Rows)
            {
                string rowData = string.Join(", ", row);
                Console.WriteLine(rowData);
            }
        }
        
        //public static void CheckPaths()
        //{
        //    string xlsFileName = "TEST_XLS.xls";
        //    string testDataFolder = "Data";
        //    string xlsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, testDataFolder, xlsFileName);

        //    Console.WriteLine("Generierter Pfad zur XLS-Datei: " + xlsFilePath);
        //    Console.WriteLine("Existiert der Pfad? " + File.Exists(xlsFilePath));

        //    if (!File.Exists(xlsFilePath))
        //    {
        //        Console.WriteLine("Die XLS-Datei existiert nicht.");
        //        Console.WriteLine("Drücke eine beliebige Taste zum Beenden...");
        //        Console.ReadKey();
        //        Environment.Exit(1);
        //    }

        //    string csvFolderName = "Results";
        //    string solutionDirectory = Directory.GetParent(Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName).FullName;
        //    string csvFolderPath = Path.Combine(solutionDirectory, "XlsXMht2Csv", "bin", "Debug", csvFolderName);
        //    string xlsFileNameWithoutExtension = Path.GetFileNameWithoutExtension(xlsFileName);
        //    string csvFileName = $"{xlsFileNameWithoutExtension}.csv";
        //    string csvFilePath = Path.Combine(csvFolderPath, csvFileName);

        //    Console.WriteLine("Generierter Pfad zur CSV-Datei: " + csvFilePath);
        //    Console.WriteLine("Existiert der Pfad? " + Directory.Exists(csvFolderPath));
        //    Console.WriteLine("Existiert die CSV-Datei? " + File.Exists(csvFilePath));

        //    Console.ReadKey();
        //}
        //public static void TestCsvPath()
        //{
        //    // Hier den hardcoded Pfad zur CSV-Datei angeben
        //    string csvFilePath = @"C:\Users\georg.schweizer\source\repos\REPO\XlsXMht2Csv_Plugin\XlsXMht2Csv\bin\Debug\Results\TEST_XLS.csv";

        //    if (!File.Exists(csvFilePath))
        //    {
        //        Console.WriteLine($"Die CSV-Datei {csvFilePath} existiert nicht.");
        //        Console.WriteLine("Drücke eine beliebige Taste zum Beenden...");
        //        Console.ReadKey();
        //        return;
        //    }

        //    Console.WriteLine("Die CSV-Datei wurde gefunden.");
        //    // Weitere Schritte zur Verarbeitung der CSV-Datei können hier hinzugefügt werden
        //    Console.ReadLine();
        //}
        //public static void ConvertXls()
        //{
        //    string xlsFileName = "TEST_XLS.xls";
        //    string testDataFolder = "Data";
        //    string xlsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, testDataFolder, xlsFileName);

        //    Console.WriteLine("Generierter Pfad zur XLS-Datei: " + xlsFilePath);
        //    Console.WriteLine("Existiert der Pfad? " + File.Exists(xlsFilePath));

        //    if (!File.Exists(xlsFilePath))
        //    {
        //        Console.WriteLine("Die XLS-Datei existiert nicht.");
        //        Console.WriteLine("Drücke eine beliebige Taste zum Beenden...");
        //        Console.ReadKey();
        //        return;
        //    }

        //    Console.WriteLine("Die XLS-Datei wurde gefunden.");

        //    FileConverter converter = new FileConverter();

        //    Console.WriteLine("Gib die Anzahl der zu überspringenden Zeilen ein:");
        //    int skipRange = int.Parse(Console.ReadLine());

        //    Console.WriteLine("Gib den Namen des Arbeitsblatts ein:");
        //    string sheetName = Console.ReadLine();

        //    Console.WriteLine("Gib das Trennzeichen (0 für Komma, 1 für Semikolon) ein:");
        //    CsvDelimiter selectedDelimiter = (CsvDelimiter)int.Parse(Console.ReadLine());
        //    Console.WriteLine($"Zeichentrennung: {selectedDelimiter.ToString()} ");
        //    string csvFolderName = "Results";
        //    string solutionDirectory = Directory.GetParent(Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName).FullName;
        //    string csvFolderPath = Path.Combine(solutionDirectory, "XlsXMht2Csv", "bin", "Debug", csvFolderName);
        //    string xlsFileNameWithoutExtension = Path.GetFileNameWithoutExtension(xlsFileName);
        //    string csvFileName = $"{xlsFileNameWithoutExtension}.csv";
        //    string csvFilePath = Path.Combine(csvFolderPath, csvFileName);
        //    Directory.CreateDirectory(csvFolderPath);

        //    converter.ConvertXlsToCsv(xlsFilePath, skipRange, sheetName, selectedDelimiter);
        //    Console.WriteLine("Konvertiere XLS-Datei in CSV...");

        //    Console.WriteLine("Convertierung abgeschlossen");
        //    Console.WriteLine($"Filepath: {csvFilePath}");
        //    Console.WriteLine($"Filepath: {csvFolderPath}");

        //    Console.ReadLine();
        //}
        //public static void ChooseSeparator()
        //{
        //    string xlsFileName = "TEST_XLS.xls";
        //    string testDataFolder = "Data";
        //    string xlsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, testDataFolder, xlsFileName);

        //    Console.WriteLine("Generierter Pfad zur XLS-Datei: " + xlsFilePath);
        //    Console.WriteLine("Existiert der Pfad? " + File.Exists(xlsFilePath));

        //    if (!File.Exists(xlsFilePath))
        //    {
        //        Console.WriteLine("Die XLS-Datei existiert nicht.");
        //        Console.WriteLine("Drücke eine beliebige Taste zum Beenden...");
        //        Console.ReadKey();
        //        return;
        //    }

        //    Console.WriteLine("Die XLS-Datei wurde gefunden.");

        //    FileConverter converter = new FileConverter();

        //    Console.WriteLine("Gib die Anzahl der zu überspringenden Zeilen ein:");
        //    int skipRange = int.Parse(Console.ReadLine());

        //    Console.WriteLine("Gib den Namen des Arbeitsblatts ein:");
        //    string sheetName = Console.ReadLine();

        //    Console.WriteLine("Wähle das Trennzeichen (0 für Komma, 1 für Semikolon):");
        //    CsvDelimiter selectedDelimiter = (CsvDelimiter)int.Parse(Console.ReadLine());

        //    Console.WriteLine("Konvertiere XLS-Datei in CSV...");

        //    Console.WriteLine("Konvertierung abgeschlossen.");

        //    string csvFolderName = "Results";
        //    string solutionDirectory = Directory.GetParent(Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName).FullName;
        //    string csvFolderPath = Path.Combine(solutionDirectory, "XlsXMht2Csv", "bin", "Debug", csvFolderName);
        //    string xlsFileNameWithoutExtension = Path.GetFileNameWithoutExtension(xlsFileName);
        //    string csvFileName = $"{xlsFileNameWithoutExtension}.csv";
        //    string csvFilePath = Path.Combine(csvFolderPath, csvFileName);

        //    FileConverter fc = new FileConverter();
        //    //string csvFilePath = Path.Combine(csvFolderPath, csvFileName);
        //    //fc.ReadCsvFileAndDisplay(csvFilePath);

        //    Console.ReadLine();

            
        //}
    }
}



    

