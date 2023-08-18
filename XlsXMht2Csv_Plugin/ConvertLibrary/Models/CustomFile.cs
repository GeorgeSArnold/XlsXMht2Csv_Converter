using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static ConvertLibrary.BusinessLogic.DelimiterManager;

namespace ConvertLibrary.Models
{
    public abstract class CustomFile
    {
        public string Id { get; set; }
        public string FileName { get; set; }
        public string FilePath { get; set; }

        

        public CustomFile(string id, string fileName, string filePath)
        {
            Id = id;
            FileName = fileName;
            FilePath = filePath;
        }
    }
}
