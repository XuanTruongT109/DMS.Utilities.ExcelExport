using DMS.Utilities.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace DMS.Utilities.ExcelExport
{
    public partial class Exporter : Form
    {
        public Exporter()
        {
            InitializeComponent();
            string[] args = new string[] { "exportconfig.xml" };
            var configPath = args[0];
            if (!File.Exists(configPath)) return;
            var serializer = new XmlSerializer(typeof(ExportCfg));
            //Export exportCfg
            //using (StreamReader rd = new StreamReader(configPath))
            //{
            //    var exportCfg = serializer.Deserialize(rd) as ExportCfg;
            //}
            //tbxConnection.Text = 
        }

        public void Export(ExportCfg cfg)
        {
            var folder = Path.GetDirectoryName(cfg.FilePath);
            if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
            var rawExporter = new RawExporter();
            if (cfg.BreakPage)
            {
                rawExporter.PageSize = cfg.PageSize;
            }
            using (var con = new SqlConnection(cfg.Connection))
            {
                con.Open();
                var cmd = new SqlCommand(cfg.Sql, con);
                cmd.CommandType = cfg.Stored ? CommandType.StoredProcedure : CommandType.Text;
                cmd.CommandTimeout = cfg.Timeout;
                var reader = cmd.ExecuteReader();
                rawExporter.WriteToExcel(cfg.FilePath, reader);
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(tbxConnection.Text))
                {
                    MessageBox.Show("Please input SQL connection string", "Required");
                    return;
                }
                if (string.IsNullOrWhiteSpace(tbxStores.Text))
                {
                    MessageBox.Show("Please input SQL stored procedure name", "Required");
                    return;
                }

                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                saveFile.Title = "Save raw data file";
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    var cfg = new ExportCfg
                    {
                        BreakPage = ckbBreak.Checked,
                        Connection = tbxConnection.Text,
                        FilePath = saveFile.FileName,
                        PageSize = (int)numPageSize.Value,
                        Sql = tbxStores.Text,
                        Stored = ckbStores.Checked,
                        Timeout = (int)tbxTimeout.Value
                    };
                    Export(cfg);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Export failed: {0}", ex.ToString()), "Error");
            }
        }
    }
}
