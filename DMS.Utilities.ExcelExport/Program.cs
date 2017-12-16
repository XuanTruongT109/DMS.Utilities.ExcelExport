using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace DMS.Utilities.ExcelExport
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            //args = new string[] { "exportconfig.xml" };
            args = new string[] { };
            if (args.Length > 0)
            {
                var configPath = args[0];
                if (!File.Exists(configPath)) return;
                var serializer = new XmlSerializer(typeof(ExportCfg));
                using(StreamReader rd = new StreamReader(configPath))
                {
                    var exportCfg = serializer.Deserialize(rd) as ExportCfg;
                    if (exportCfg == null) return;
                    exportCfg.FilePath = exportCfg.FilePath.Replace("[NOW]", DateTime.Now.ToString("yyyyMMddhhmmssfff"));
                    new Exporter().Export(exportCfg);
                }
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Exporter());
            }
        }
    }
}
