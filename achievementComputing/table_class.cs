using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Drawing;

namespace achievementComputing
{
    class table_class
    {
        internal int startRow;
        internal int startCol;
        internal int endRow;
        internal int endCol;
		internal int width;
		internal int height;
		internal ExcelRange range;
        internal table_class() { }
		//internal table_class(ExcelRange rg)
		//{
		//    range = rg;
		//    startRow = range.Start.Row;
		//    startCol = range.Start.Column;
		//    endRow = range.End.Row;
		//    endCol = range.End.Column;
		//    //computeWidthAndHeight();
		//}

		internal void computeWidthAndHeight()
		{
			width = endCol - startCol + 1;
			height = endRow - startRow + 1;
		}
		internal ExcelRange recaptureRange(ExcelWorksheet sheet)
        {
            range = sheet.Cells[startRow, startCol, endRow, endCol];
            return (range);
        }

        internal void colour(Color color)
        {
			range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
			range.Style.Fill.BackgroundColor.SetColor(color);
        }
        internal void Border()
		{
            range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
        }
    }
}
