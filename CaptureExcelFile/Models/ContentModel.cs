using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaptureExcelFile.Models
{
    public class ContentModel
    {
        public string ProductName { get; set; }
        public int StockStartMonth { get; set; }
        public int ImportVK { get; set; }
        public int ImportNCQ { get; set; }
        public int NhatNam { get; set; }
        public int ImportSW { get; set; }
        public int ImportCLK { get; set; }
        public int ImportTL { get; set; }
        public int ChangeShield { get; set; }
        public int ExportSold { get; set; }
        public int ExportTransport { get; set; }
        public int StockEndMonth { get; set; }
        public int MiniStock { get; set; }
        public int Different { get; set; }  
        public string OldDescription { get; set; }

    }
}
