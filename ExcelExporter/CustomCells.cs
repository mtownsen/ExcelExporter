//CustomCells

using System;
using DocumentFormat.OpenXml.Spreadsheet;
namespace CreateExcelSpreadsheet
{
    public class TextCell : Cell
    {
        public TextCell(string header, string text, int index)
        {
            this.DataType = CellValues.InlineString;
            this.CellReference = header + index;
            //Add text to the text cell.
            this.InlineString = new InlineString { Text = new Text { Text = text } };
        }
    }
    public class NumberCell : Cell
    {
        public NumberCell(string header, string text, int index)
        {
            this.DataType = CellValues.Number;
            this.CellReference = header + index;
            this.CellValue = new CellValue(text);
        }
    }
    public class FormatedNumberCell : NumberCell
    {
        public FormatedNumberCell(string header, string text, int index)
            : base(header, text, index)
        {
            this.StyleIndex = 2;
        }
    }
    public class DateCell : Cell
    {
        public DateCell(string header, DateTime dateTime, int index)
        {
            this.DataType = CellValues.Date;
            this.CellReference = header + index;
            this.StyleIndex = 1;
            this.CellValue = new CellValue { Text = dateTime.ToOADate().ToString() }; ;
        }
    }
    public class FomulaCell : Cell
    {
        public FomulaCell(string header, string text, int index)
        {
            this.CellFormula = new CellFormula { CalculateCell = true, Text = text };
            this.DataType = CellValues.Number;
            this.CellReference = header + index;
            this.StyleIndex = 2;
        }
    }
    public class HeaderCell : TextCell
    {
        public HeaderCell(string header, string text, int index) :
            base(header, text, index)
        {
            this.StyleIndex = 11;
        }
    }
}