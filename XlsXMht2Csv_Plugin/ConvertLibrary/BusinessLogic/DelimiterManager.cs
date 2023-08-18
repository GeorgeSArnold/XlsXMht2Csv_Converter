using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConvertLibrary.BusinessLogic
{
    public class DelimiterManager
    {
        public enum CsvDelimiter
        {
            Comma,
            Semicolon
        }
        public static char GetDelimiterChar(CsvDelimiter delimiter)
        {
            switch (delimiter)
            {
                case CsvDelimiter.Comma:
                    return ',';
                case CsvDelimiter.Semicolon:
                    return ';';
                default:
                    throw new ArgumentException("Unsupported delimiter");
            }
        }
    }
}
