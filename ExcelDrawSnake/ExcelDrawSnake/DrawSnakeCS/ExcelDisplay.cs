using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;
namespace ExcelDrawSnake.DrawSnakeCS
{
    static class ExcelDisplay
    {
        /// <summary>
        /// 刷新显示
        /// </summary>
        /// <param name="xlWorksheet">要绘制的worksheet</param>
        /// <param name="snakePointList">蛇身体的坐标集合</param>
        /// <param name="color">蛇身体的颜色</param>
        public static void Display(Excel.Worksheet xlWorksheet, List<Point> snakePointList,Color color)
        {
            try
            {
                var lastPoint = snakePointList[0];
                Excel.Range lsatRange = xlWorksheet.Range[xlWorksheet.Cells[lastPoint.X + 1, lastPoint.Y + 1], xlWorksheet.Cells[lastPoint.X + 1, lastPoint.Y + 1]];

                snakePointList.RemoveAt(0);
                //避免当蛇穿过自己的身体时使身体变空
                if (!snakePointList.Any(point => (point.X == lastPoint.X) && (point.Y == lastPoint.Y)))
                {
                    lsatRange.Interior.Color = Color.White;
                }
                snakePointList.Insert(0, lastPoint);

                var currentPoint = snakePointList.Last();
                Excel.Range currentRange = xlWorksheet.Range[xlWorksheet.Cells[currentPoint.X + 1, currentPoint.Y + 1], xlWorksheet.Cells[currentPoint.X + 1, currentPoint.Y + 1]];
                currentRange.Interior.Color = color;

                Debug.WriteLine(" 蛇头 X:{0} Y:{1}", currentPoint.X, currentPoint.Y);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        /// <summary>
        /// 显示初始化的蛇
        /// </summary>
        /// <param name="xlWorksheet">要绘制的worksheet</param>
        /// <param name="snakePointList">蛇身体的坐标集合</param>
        /// <param name="color">初始化时蛇的颜色</param>
        public static void DisplayInit(Excel.Worksheet xlWorksheet,List<Point> snakePointList,Color color)
        {
            foreach (var point in snakePointList)
            {
                Excel.Range range = xlWorksheet.Range[xlWorksheet.Cells[point.X + 1, point.Y + 1], xlWorksheet.Cells[point.X + 1, point.Y + 1]];
                range.Interior.Color = color;

                Debug.WriteLine("初始化的点 X:{0} Y:{1}", point.X, point.Y);
            }
        }

        /// <summary>
        /// 显示果实
        /// </summary>
        /// <param name="xlWorksheet"></param>
        /// <param name="point"></param>
        /// <param name="color">果实的颜色</param>
        public static void DislayRandomPoint(Excel.Worksheet xlWorksheet, Point point, Color color)
        {
            try
            {
                Excel.Range range = xlWorksheet.Range[xlWorksheet.Cells[point.X + 1, point.Y + 1], xlWorksheet.Cells[point.X + 1, point.Y + 1]];
                range.Interior.Color = color;

                Debug.WriteLine("       果实 X:{0} Y:{1}", point.X, point.Y);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }
    }
}
