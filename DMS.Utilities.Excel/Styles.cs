using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.Utilities.Excel
{
    public class StyleManager
    {
        const uint BASE_NUMBER_FORMAT_ID = 166;
        private List<xBorder> borders;
        private List<xFill> fills;
        private List<xFont> fonts;
        private List<xCellFormat> cellFormats;
        private List<xNumberFormat> numFormats;

        public Fonts GetXmlFonts()
        {
            var xmlFonts = new Fonts();
            for (int i = 0; i < fonts.Count; i++)
            {
                xmlFonts.Append(fonts[i].ToXmlStyle());
            }
            xmlFonts.Count = (uint)fonts.Count;
            xmlFonts.KnownFonts = true;
            return xmlFonts;
        }

        public Fills GetXmlFills()
        {
            var xmlFills = new Fills();
            for (int i = 0; i < fills.Count; i++)
            {
                xmlFills.Append(fills[i].ToXmlStyle());
            }
            xmlFills.Count = (uint)fills.Count;
            return xmlFills;
        }

        public Borders GetXmlBorders()
        {
            var xmlBorders = new Borders();
            for (int i = 0; i < borders.Count; i++)
            {
                xmlBorders.Append(borders[i].ToXmlStyle());
            }
            xmlBorders.Count = (uint)borders.Count;
            return xmlBorders;
        }

        public CellFormats GetXmlCellFormats()
        {
            var xmlCellFormats = new CellFormats();
            for (int i = 0; i < cellFormats.Count; i++)
            {
                xmlCellFormats.Append(cellFormats[i].ToXmlStyle());
            }
            xmlCellFormats.Count = (uint)cellFormats.Count;
            return xmlCellFormats;
        }

        public NumberingFormats GetXmlNumberFormats()
        {
            var xmlNumberingFmt = new NumberingFormats();
            for (int i = 0; i < numFormats.Count; i++)
            {
                xmlNumberingFmt.Append(numFormats[i].ToXmlStyle((uint)i + BASE_NUMBER_FORMAT_ID));
            }
            xmlNumberingFmt.Count = (uint)numFormats.Count;
            return xmlNumberingFmt;
        }

        public StyleManager()
        {
            borders = new List<xBorder>();
            fills = new List<xFill>();
            fonts = new List<xFont>();
            cellFormats = new List<xCellFormat>();
            numFormats = new List<xNumberFormat>();
            InitDefault();
        }

        public virtual void InitDefault()
        {
            fonts.Add(new xFont()
            {
                Name = "Arial",
                Size = 11,
                Color = "000000"
            });


            fills.Add(new xFill()
            {
                PatternType = PatternValues.None
            });
            fills.Add(new xFill()
            {
                PatternType = PatternValues.Gray125
            });
            borders.Add(new xBorder());
            cellFormats = new List<xCellFormat>() { new xCellFormat() { BorderId = 0, FillId = 0, FontId = 0 } };
        }

        public uint GetXmlId(xNumberFormat numberFormat)
        {
            for (int i = 0; i < numFormats.Count; i++)
            {
                if (numFormats[i].Equals(numberFormat)) return (uint)i + BASE_NUMBER_FORMAT_ID;
            }
            numFormats.Add(numberFormat);
            return ((uint)numFormats.Count + BASE_NUMBER_FORMAT_ID - 1);
        }

        public uint GetXmlId(xBorder border)
        {
            for (int i = 0; i < borders.Count; i++)
            {
                if (borders[i].Equals(border)) return (uint)i;
            }
            borders.Add(border);
            return (uint)(borders.Count - 1);
        }

        public uint GetXmlId(xFill fill)
        {
            for (int i = 0; i < fills.Count; i++)
            {
                if (fills[i].Equals(fill)) return (uint)i;
            }
            fills.Add(fill);
            return (uint)(fills.Count - 1);
        }

        public uint GetXmlId(xFont font)
        {
            for (int i = 0; i < fonts.Count; i++)
            {
                if (fonts[i].Equals(font)) return (uint)i;
            }
            fonts.Add(font);
            return (uint)(fonts.Count - 1);
        }
        public uint GetXmlId(xCellFormat cf)
        {
            for (int i = 0; i < cellFormats.Count; i++)
            {
                if (cellFormats[i].Equals(cf)) return (uint)i;
            }
            cellFormats.Add(cf);
            return (uint)(cellFormats.Count - 1);
        }
    }

    public class xColumn : IEquatable<xColumn>
    {
        public double? Width { get; set; }

        public uint FromColumn { get; set; }
        public uint ToColumn { get; set; }

        public Column ToXmlStyle()
        {
            var col = new Column() { Min = FromColumn, Max = ToColumn };

            if (Width != null)
            {
                col.CustomWidth = true;
                col.Width = Width.Value;
            }
            return col;
        }

        public bool Equals(xColumn other)
        {
            return Width == other.Width
                && FromColumn == other.FromColumn
                && ToColumn == other.ToColumn;
        }
    }

    public class xNumberFormat : IEquatable<xNumberFormat>
    {
        public string FormatCode { get; set; }
        public NumberingFormat ToXmlStyle(uint numberFormatId)
        {
            var numFmt = new NumberingFormat();
            numFmt.FormatCode = FormatCode;
            numFmt.NumberFormatId = numberFormatId;
            return numFmt;
        }
        public bool Equals(xNumberFormat other)
        {
            return FormatCode == other.FormatCode;
        }
    }

    public class xCellFormat : IEquatable<xCellFormat>
    {
        public uint FontId { get; set; }
        public uint BorderId { get; set; }
        public uint FillId { get; set; }
        public uint NumberFormatId { get; set; }
        public string HAlign { get; set; }
        public string VAlign { get; set; }

        public bool WrapText { get; set; }
        public bool ShrinkToFit { get; set; }

        public CellFormat ToXmlStyle()
        {
            var cf = new CellFormat();
            cf.FontId = FontId;
            cf.BorderId = BorderId;
            cf.FillId = FillId;
            var align = new Alignment();
            if (FontId > 0) cf.ApplyFont = true;
            if (BorderId > 0) cf.ApplyBorder = true;
            if (FillId > 0) cf.ApplyFill = true;
            if (NumberFormatId > 0)
                cf.NumberFormatId = NumberFormatId;
            if (WrapText)
                align.WrapText = WrapText;
            if (ShrinkToFit)
                align.ShrinkToFit = ShrinkToFit;
            if (HAlign != null)
            {
                align.Horizontal = (HorizontalAlignmentValues)Enum.Parse(typeof(HorizontalAlignmentValues), HAlign);
            }
            if (VAlign != null)
            {
                align.Vertical = (VerticalAlignmentValues)Enum.Parse(typeof(VerticalAlignmentValues), VAlign);
            }

            if (HAlign != null || VAlign != null)
                cf.Alignment = align;
            return cf;
        }

        public bool Equals(xCellFormat other)
        {
            return FontId == other.FontId
                && BorderId == other.BorderId
                && FillId == other.FillId
                && HAlign == other.HAlign
                && VAlign == other.VAlign
                && NumberFormatId == other.NumberFormatId
                && WrapText == other.WrapText
                && ShrinkToFit == other.ShrinkToFit;
        }
    }

    public class xFill : IEquatable<xFill>
    {
        public PatternValues PatternType { get; set; }
        public string FillCollor { get; set; }

        public Fill ToXmlStyle()
        {
            var fill = new Fill();
            var pattern = new PatternFill() { PatternType = PatternType };
            if (!string.IsNullOrEmpty(FillCollor))
                pattern.ForegroundColor = new ForegroundColor() { Rgb = HexBinaryValue.FromString(FillCollor) };
            fill.Append(pattern);
            return fill;
        }

        public bool Equals(xFill other)
        {
            return PatternType == other.PatternType
                && FillCollor == other.FillCollor;
        }
    }

    public class xBorder : IEquatable<xBorder>
    {
        public bool All
        {
            set
            {
                Left = Right = Top = Bottom = value;
            }
        }
        public bool Left { get; set; }
        public bool Right { get; set; }
        public bool Top { get; set; }
        public bool Bottom { get; set; }
        public bool Diagonal { get; set; }
        public string Color
        {
            get { return color.Rgb.Value; }
            set
            {
                color = new Color { Rgb = value };
            }
        }
        public int Style
        {
            get { return (int)style; }
            set { style = (BorderStyleValues)value; }
        }

        private BorderStyleValues style;
        private Color color;
        public xBorder()
        {
            color = new Color { Auto = true };
            style = BorderStyleValues.Thin;
        }

        public Border ToXmlStyle()
        {
            var border = new Border();
            if (Left) border.Append(new LeftBorder() { Color = (Color)color.Clone(), Style = style });
            if (Right) border.Append(new RightBorder() { Color = (Color)color.Clone(), Style = style });
            if (Top) border.Append(new TopBorder() { Color = (Color)color.Clone(), Style = style });
            if (Bottom) border.Append(new BottomBorder() { Color = (Color)color.Clone(), Style = style });
            if (Diagonal) border.Append(new DiagonalBorder() { Color = (Color)color.Clone(), Style = style });
            return border;
        }

        public bool Equals(xBorder other)
        {
            return this.Left == other.Left
                && this.Right == other.Right
                && this.Top == other.Top
                && this.Bottom == other.Bottom
                && this.Diagonal == other.Diagonal
                && this.color?.Rgb?.Value == other.color?.Rgb?.Value;
        }
    }

    public class xFont : IEquatable<xFont>
    {
        public int Family { get; set; }
        public string Name { get; set; }
        public double Size { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }

        public string Color
        {
            get { return color?.Rgb?.Value; }
            set
            {
                color = new Color { Rgb = value };
            }
        }
        private Color color = new Color { Auto = true };

        public Font ToXmlStyle()
        {
            var font = new Font();
            if (!string.IsNullOrEmpty(Name)) font.Append(new FontName() { Val = Name });
            if (Size > 0) font.Append(new FontSize() { Val = Size });
            if (color != null) font.Append(color);
            if (Bold) font.Append(new Bold());
            if (Italic) font.Append(new Italic());
            if (Underline) font.Append(new Underline());
            if (Family > 0) font.Append(new FontFamily() { Val = Family });

            return font;
        }

        public bool Equals(xFont other)
        {
            return Name == other.Name
                && Size == other.Size
                && color?.Rgb?.Value == other.color?.Rgb?.Value
                && Bold == other.Bold
                && Italic == other.Italic
                && Underline == other.Underline
                && Family == other.Family;

        }
    }
}
