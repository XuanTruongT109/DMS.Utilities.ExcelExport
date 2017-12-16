using DMS.Utilities.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DMS.Utilities.Excel
{
    public class RawExporter : ExcelWriter
    {
        private int rowCount;
        private bool hasNext;
        public int PageSize { get; set; } = 0;
        private Dictionary<int, Func<uint, string, Cell>> cellFactory;

        public override List<string> GetMergeCells()
        {
            return new List<string>();
        }

        protected override bool NextSheet()
        {
            return hasNext;
        }

        public override void WriteData(OpenXmlWriter oxw, IDataReader reader, StyleManager styleMgr)
        {
            var numFmtId = styleMgr.GetXmlId(new xNumberFormat { FormatCode = "#,##0" });
            var textFontId = styleMgr.GetXmlId(new xFont()
            {
                Name = "Arial",
                Family = 2,
                Size = 8
            });
            var bodyBorderId = styleMgr.GetXmlId(new xBorder()
            {
                Left = true,
                Right = true,
                Top = true,
                Bottom = true
            });

            var textStyleId = styleMgr.GetXmlId(new xCellFormat() { FontId = textFontId, HAlign = "Left", BorderId = bodyBorderId, WrapText = true, VAlign = "Center" });
            var numberStyleId = styleMgr.GetXmlId(new xCellFormat() { FontId = textFontId, HAlign = "Right", BorderId = bodyBorderId, WrapText = true, NumberFormatId = numFmtId, VAlign = "Center" });

            hasNext = false;
            uint idx = 2;
            while (reader.Read())
            {
                var row = new Row() { RowIndex = idx };
                for (int i = 0; i < reader.FieldCount; i++)
                {
                    var cell = cellFactory[i](idx, reader.GetValue(i)?.ToString() ?? string.Empty);
                    row.Append(cell);
                }
                oxw.WriteElement(row);
                idx++;
                rowCount++;
                if (PageSize != 0 && rowCount >= PageSize)
                {
                    hasNext = true;
                    break;
                }
            }
        }

        private bool IsNumericType(Type type)
        {
            return new Type[] { typeof(int), typeof(float), typeof(double), typeof(Int64), typeof(Decimal) }.Contains(type);
        }

        public override void WriteHeader(OpenXmlWriter oxw, IDataReader reader, StyleManager styleMgr)
        {
            rowCount = 0;
            hasNext = true;
            var fontId = styleMgr.GetXmlId(new xFont()
            {
                Name = "Arial",
                Family = 2,
                Size = 10,
                Bold = true
            });
            var borderId = styleMgr.GetXmlId(new xBorder()
            {
                Left = true,
                Right = true,
                Top = true,
                Bottom = true
            });
            var fillId = styleMgr.GetXmlId(new xFill()
            {
                PatternType = PatternValues.Solid,
                FillCollor = "BFBFBF"
            });

            var styleId = styleMgr.GetXmlId(new xCellFormat() { FontId = fontId, HAlign = "Center", VAlign = "Center", BorderId = borderId, FillId = fillId, WrapText = true });

            cellFactory = new Dictionary<int, Func<uint, string, Cell>>();
            var numFmtId = styleMgr.GetXmlId(new xNumberFormat { FormatCode = "#,##0" });
            var textFontId = styleMgr.GetXmlId(new xFont()
            {
                Name = "Arial",
                Family = 2,
                Size = 8
            });
            var bodyBorderId = styleMgr.GetXmlId(new xBorder()
            {
                Left = true,
                Right = true,
                Top = true,
                Bottom = true
            });

            var textStyleId = styleMgr.GetXmlId(new xCellFormat() { FontId = textFontId, HAlign = "Left", BorderId = bodyBorderId, WrapText = true, VAlign = "Center" });
            var numberStyleId = styleMgr.GetXmlId(new xCellFormat() { FontId = textFontId, HAlign = "Right", BorderId = bodyBorderId, WrapText = true, NumberFormatId = numFmtId, VAlign = "Center" });


            var row = new Row() { RowIndex = 1 };
            for (int i = 0; i < reader.FieldCount; i++)
            {
                var colname = xCellAddress.GetColumnName((uint)i + 1);
                var cell = new Cell()
                {
                    CellReference = colname + "1",
                    StyleIndex = styleId,
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString() { Text = new Text(reader.GetName(i)) }
                };
                row.Append(cell);
                if (IsNumericType(reader.GetFieldType(i)))
                {
                    var func = new Func<uint, string, Cell>((idx, value) =>
                    {
                        return new Cell()
                        {
                            CellReference = colname + idx,
                            StyleIndex = numberStyleId,
                            DataType = CellValues.Number,
                            CellValue = new CellValue() { Text = value }
                        };
                    });
                    cellFactory.Add(i, func);
                }
                else
                {
                    var func = new Func<uint, string, Cell>((idx, value) =>
                    {
                        return new Cell()
                        {
                            CellReference = colname + idx,
                            StyleIndex = textStyleId,
                            DataType = CellValues.InlineString,
                            InlineString = new InlineString() { Text = new Text(value) }
                        };
                    });
                    cellFactory.Add(i, func);
                }

            }
            oxw.WriteElement(row);



        }
    }
}
