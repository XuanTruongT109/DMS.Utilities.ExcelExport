using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.Utilities.Excel
{
    public class xCellAddress
    {
        public uint Row { get; set; }
        public uint Column { get; set; }

        public string ToCellRef()
        {
            return string.Format("{0}{1}", GetColumnName(Column), Row);
        }

        public static uint GetColumnIndex(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return (uint)sum;
        }

        public static string GetColumnName(uint columnNumber)
        {
            uint dividend = columnNumber;
            string columnName = String.Empty;
            uint modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (uint)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }

    public class xCellData
    {
        public xCellAddress Address { get; set; }
        public string Value { get; set; }
    }

    public abstract class ExcelWriter
    {
        public abstract void WriteHeader(OpenXmlWriter oxw, IDataReader reader, StyleManager styleMgr);
        public abstract void WriteData(OpenXmlWriter oxw, IDataReader reader, StyleManager styleMgr);
        protected virtual bool NextSheet() { return false; }
        public virtual List<xColumn> GetColumnFormat()
        {
            return new List<xColumn>();
        }
        public abstract List<string> GetMergeCells();


        public void WriteToExcel(string filename, IDataReader reader)
        {
            using (SpreadsheetDocument xl = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                xl.AddWorkbookPart();
                WorkbookPart workbookPart = xl.WorkbookPart;
                StyleManager styleMgr = new StyleManager();
                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();

                #region Create worksheet
                var sheets = WriteSheet(workbookPart, reader, styleMgr, stylePart);
                #endregion Create worksheet

                var oxw = OpenXmlWriter.Create(xl.WorkbookPart);
                oxw.WriteStartElement(new Workbook());
                oxw.WriteStartElement(new Sheets());

                // you can use object initialisers like this only when the properties
                // are actual properties. SDK classes sometimes have property-like properties
                // but are actually classes. For example, the Cell class has the CellValue
                // "property" but is actually a child class internally.
                // If the properties correspond to actual XML attributes, then you're fine.
                foreach (var sheet in sheets)
                {
                    oxw.WriteElement(sheet);
                }

                // this is for Sheets
                oxw.WriteEndElement();
                // this is for Workbook
                oxw.WriteEndElement();
                oxw.Close();

                WriteStyle(stylePart, styleMgr);

                xl.Close();
            }
        }

        protected virtual List<Sheet> WriteSheet(WorkbookPart workbookPart, IDataReader reader, StyleManager styleMgr, WorkbookStylesPart stylePart)
        {
            uint sheetId = 1;
            var sheets = new List<Sheet>();
            do
            {
                WorksheetPart wsp = workbookPart.AddNewPart<WorksheetPart>();
                var oxw = OpenXmlWriter.Create(wsp);
                oxw.WriteStartElement(new Worksheet());

                var columns = GetColumnFormat().Select(x => x.ToXmlStyle()).ToArray();
                if (columns.Length > 0)
                {
                    var colFmt = new Columns(columns);
                    oxw.WriteElement(colFmt);
                }

                oxw.WriteStartElement(new SheetData());

                WriteHeader(oxw, reader, styleMgr);
                WriteData(oxw, reader, styleMgr);

                // this is for SheetData
                oxw.WriteEndElement();

                var mergeCells = GetMergeCells()
                    .Select(x => new MergeCell { Reference = new StringValue(x) })
                    .ToArray();
                if (mergeCells.Length > 0)
                {
                    oxw.WriteElement(new MergeCells(mergeCells) { Count = (uint)mergeCells.Length });
                }
                // this is for Worksheet
                oxw.WriteEndElement();
                oxw.Close();
                sheets.Add(new Sheet()
                {
                    Name = string.Format("Data{0}", sheetId),
                    SheetId = sheetId,
                    Id = workbookPart.GetIdOfPart(wsp)
                });
                sheetId++;
            } while (NextSheet());
            return sheets;
        }

        private void WriteStyle(WorkbookStylesPart stylePart, StyleManager styleMgr)
        {
            var borders = styleMgr.GetXmlBorders();
            var fonts = styleMgr.GetXmlFonts();
            var fills = styleMgr.GetXmlFills();
            var cfms = styleMgr.GetXmlCellFormats();
            var numFmts = styleMgr.GetXmlNumberFormats();

            Stylesheet stylesheet = new Stylesheet(numFmts, fonts, fills, borders, cfms);
            var oxw = OpenXmlWriter.Create(stylePart);
            oxw.WriteElement(stylesheet);
            oxw.Close();
        }
    }
}
