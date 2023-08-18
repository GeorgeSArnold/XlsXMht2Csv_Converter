using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ConvertLibrary.BusinessLogic.DelimiterManager;

namespace ConvertLibrary.Models
{
    public class Csv_File : CustomFile
    {
        public Csv_File(string id, string fileName, string filePath) : base(id, fileName, filePath) { }
    }
}
