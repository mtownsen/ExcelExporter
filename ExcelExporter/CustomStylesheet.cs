//CustomStylesheet.cs
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
namespace CreateExcelSpreadsheet
{
    public class CustomStylesheet : Stylesheet
    {
        public CustomStylesheet()
        {
            var fonts = new Fonts();
            var font = new DocumentFormat.OpenXml.Spreadsheet.Font();
            var fontName = new FontName { Val = StringValue.FromString("Arial") };
            var fontSize = new FontSize { Val = DoubleValue.FromDouble(11) };
            font.FontName = fontName;
            font.FontSize = fontSize;
            fonts.Append(font);
            //Font Index 1
            font = new DocumentFormat.OpenXml.Spreadsheet.Font();
            fontName = new FontName { Val = StringValue.FromString("Arial") };
            fontSize = new FontSize { Val = DoubleValue.FromDouble(12) };
            font.FontName = fontName;
            font.FontSize = fontSize;
            font.Bold = new Bold();
            fonts.Append(font);
            fonts.Count = UInt32Value.FromUInt32((uint)fonts.ChildElements.Count);
            var fills = new Fills();
            var fill = new Fill();
            var patternFill = new PatternFill { PatternType = PatternValues.None };
            fill.PatternFill = patternFill;
            fills.Append(fill);
            fill = new Fill();
            patternFill = new PatternFill { PatternType = PatternValues.Gray125 };
            fill.PatternFill = patternFill;
            fills.Append(fill);
            //Fill index  2
            fill = new Fill();
            patternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor()
            };
            patternFill.ForegroundColor =
               TranslateForeground(System.Drawing.Color.FromArgb(26,94,144));
            patternFill.BackgroundColor =
                new BackgroundColor { Rgb = patternFill.ForegroundColor.Rgb };
            fill.PatternFill = patternFill;
            fills.Append(fill);
            //Fill index  3
            fill = new Fill();
            patternFill = new PatternFill
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor()
            };
            patternFill.ForegroundColor =
               TranslateForeground(System.Drawing.Color.LightGray);
            patternFill.BackgroundColor =
               new BackgroundColor { Rgb = patternFill.ForegroundColor.Rgb };
            fill.PatternFill = patternFill;
            fills.Append(fill);
            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);
            var borders = new Borders();
            var border = new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);
            //All Boarder Index 1
            border = new Border
            {
                LeftBorder = new LeftBorder { Style = BorderStyleValues.Thin },
                RightBorder = new RightBorder { Style = BorderStyleValues.Thin },
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);
            //Top and Bottom Boarder Index 2
            border = new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder { Style = BorderStyleValues.Thin },
                BottomBorder = new BottomBorder { Style = BorderStyleValues.Thin },
                DiagonalBorder = new DiagonalBorder()
            };
            borders.Append(border);
            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);
            var cellStyleFormats = new CellStyleFormats();
            var cellFormat = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0
            };
            cellStyleFormats.Append(cellFormat);
            cellStyleFormats.Count =
               UInt32Value.FromUInt32((uint)cellStyleFormats.ChildElements.Count);
            uint iExcelIndex = 164;
            var numberingFormats = new NumberingFormats();
            var cellFormats = new CellFormats();
            cellFormat = new CellFormat
            {
                NumberFormatId = 0,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0
            };
            cellFormats.Append(cellFormat);
            var nformatDateTime = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss")
            };
            numberingFormats.Append(nformatDateTime);
            var nformat4Decimal = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("#,##0.0000")
            };
            numberingFormats.Append(nformat4Decimal);
            var nformat2Decimal = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++),
                FormatCode = StringValue.FromString("#,##0.00")
            };
            numberingFormats.Append(nformat2Decimal);
            var nformatForcedText = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(iExcelIndex),
                FormatCode = StringValue.FromString("@")
            };
            numberingFormats.Append(nformatForcedText);
            // index 1
            // Cell Standard Date format 
            cellFormat = new CellFormat
            {
                NumberFormatId = 14,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);
            // Index 2
            // Cell Standard Number format with 2 decimal placing
            cellFormat = new CellFormat
            {
                NumberFormatId = 4,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);
            // Index 3
            // Cell Date time custom format
            cellFormat = new CellFormat
            {
                NumberFormatId = nformatDateTime.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);
            // Index 4
            // Cell 4 decimal custom format
            cellFormat = new CellFormat
            {
                NumberFormatId = nformat4Decimal.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);
            // Index 5
            // Cell 2 decimal custom format
            cellFormat = new CellFormat
            {
                NumberFormatId = nformat2Decimal.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);
            // Index 6
            // Cell forced number text custom format
            cellFormat = new CellFormat
            {
                NumberFormatId = nformatForcedText.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);
            // Index 7
            // Cell text with font 12 
            cellFormat = new CellFormat
            {
                NumberFormatId = nformatForcedText.NumberFormatId,
                FontId = 1,
                FillId = 0,
                BorderId = 0,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);
            // Index 8
            // Cell text
            cellFormat = new CellFormat
            {
                NumberFormatId = nformatForcedText.NumberFormatId,
                FontId = 0,
                FillId = 0,
                BorderId = 1,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);
            // Index 9
            // Coloured 2 decimal cell text
            cellFormat = new CellFormat
            {
                NumberFormatId = nformat2Decimal.NumberFormatId,
                FontId = 0,
                FillId = 2,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);
            // Index 10
            // Coloured cell text
            cellFormat = new CellFormat
            {
                NumberFormatId = nformatForcedText.NumberFormatId,
                FontId = 0,
                FillId = 2,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);
            // Index 11
            // Coloured cell text
            cellFormat = new CellFormat
            {
                NumberFormatId = nformatForcedText.NumberFormatId,
                FontId = 1,
                FillId = 3,
                BorderId = 2,
                FormatId = 0,
                ApplyNumberFormat = BooleanValue.FromBoolean(true)
            };
            cellFormats.Append(cellFormat);
            numberingFormats.Count =
              UInt32Value.FromUInt32((uint)numberingFormats.ChildElements.Count);
            cellFormats.Count = UInt32Value.FromUInt32((uint)cellFormats.ChildElements.Count);
            this.Append(numberingFormats);
            this.Append(fonts);
            this.Append(fills);
            this.Append(borders);
            this.Append(cellStyleFormats);
            this.Append(cellFormats);
            var css = new CellStyles();
            var cs = new CellStyle
            {
                Name = StringValue.FromString("Normal"),
                FormatId = 0,
                BuiltinId = 0
            };
            css.Append(cs);
            css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
            this.Append(css);
            var dfs = new DifferentialFormats { Count = 0 };
            this.Append(dfs);
            var tss = new TableStyles
            {
                Count = 0,
                DefaultTableStyle = StringValue.FromString("TableStyleMedium9"),
                DefaultPivotStyle = StringValue.FromString("PivotStyleLight16")
            };
            this.Append(tss);
        }
        private static ForegroundColor TranslateForeground(System.Drawing.Color fillColor)
        {
            return new ForegroundColor()
            {
                Rgb = new HexBinaryValue()
                {
                    Value =
                        System.Drawing.ColorTranslator.ToHtml(
                        System.Drawing.Color.FromArgb(
                            fillColor.A,
                            fillColor.R,
                            fillColor.G,
                            fillColor.B)).Replace("#", "")
                }
            };
        }
    }
}