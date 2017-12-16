using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.Utilities.ExcelExport
{
    public class ExportCfg
    {
        public string Sql { get; set; }
        public string Connection { get; set; }
        public bool Stored { get; set; }
        public string FilePath { get; set; }
        public bool BreakPage { get; set; }
        public int PageSize { get; set; }
        public int Timeout { get; set; }
    }
}
