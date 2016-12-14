using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDrawSnake.DrawSnakeCS
{
    public class ExcelSet
    {
        private static Excel.Application _xlApp;

        /// <summary>
        /// 设置贪吃蛇的活动范围大小
        /// </summary>
        /// <param name="xlWorksheet"></param>
        /// <param name="sizeHeightX">活动范围的高就是行数</param>
        /// <param name="sSizeWidthY">活动范围的宽就是列数</param>
        /// <param name="rowHeight">小方块的高</param>
        /// <param name="columnWidth">小方块的宽</param>
        public static void SetCellSize(Excel.Worksheet xlWorksheet, int sizeHeightX, int sSizeWidthY, double rowHeight,
            double columnWidth)
        {
            _xlApp = xlWorksheet.Application;
            Excel.Range range = xlWorksheet.Range[xlWorksheet.Cells[1, 1], xlWorksheet.Cells[sizeHeightX, sSizeWidthY]];
            _xlApp.ActiveWindow.DisplayGridlines = false; //去掉网格线 

            range.Interior.Color = Color.White;

            //设置行高列宽
            range.RowHeight = rowHeight;
            range.ColumnWidth = columnWidth;
            //range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;//每个单元格都有边框
            //range的外边框
            range.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium,
                Excel.XlColorIndex.xlColorIndexAutomatic, System.Drawing.Color.Black.ToArgb());
        }
    }
}