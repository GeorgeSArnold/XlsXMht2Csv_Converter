using System.Collections.Generic;
using static ConvertLibrary.BusinessLogic.DelimiterManager;

namespace ConvertLibrary.Models
{
    public class Table
    {
        public List<string[]> Rows { get; set; }
        public List<string[]> Columns { get; set; } //TODO:Festlegung/get! der Columns 
        public string SheetName { get; set; } // Sheet Bezeichnung
        public int SheetSet { get; set; } // Menge der Sheets
        public int SkipRange { get; set; } // Menge der Ausgelassenen Felder

        public CsvDelimiter Delimiter { get; set; } // Trennzeichen-Auswahl

        public Table(List<string[]> rows, List<string[]> columns, string sheetName,int sheetSet,int skipRange, CsvDelimiter delimiter)
        {
            Rows = rows;
            SheetName = sheetName;
            SheetSet = sheetSet;
            SkipRange = skipRange;
            Delimiter = delimiter;
        }
    }
}

