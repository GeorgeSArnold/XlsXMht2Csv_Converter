using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ConvertLibrary.BusinessLogic.DelimiterManager;

namespace ConvertLibrary.Models
{
    public class Xlsx_File : Xls_File
    {
        public Xlsx_File(string id, string fileName, string filePath) : base(id, fileName, filePath) { }
    }
}
