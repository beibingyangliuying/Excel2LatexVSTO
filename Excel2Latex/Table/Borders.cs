using static Excel2Latex.Utilities;
using Excel = Microsoft.Office.Interop.Excel;
using Index = Microsoft.Office.Interop.Excel.XlBordersIndex;

namespace Excel2Latex.Table
{
    internal class Borders
    {
        public bool[,] HorizontalBorders { get; }
        public bool[,] VerticalBorders { get; }
        public Borders(Excel.Range range)
        {
            var m = range.Rows.Count;
            var n = range.Columns.Count;
            HorizontalBorders = new bool[m + 1, n];
            VerticalBorders = new bool[m, n + 1];

            for (var i = 0; i < m; i++)
            {
                for (var j = 0; j < n; j++)
                {
                    var cell = range.Item[i + 1, j + 1];
                    var left = IfBorder((int)cell.Borders.Item[Index.xlEdgeLeft].LineStyle);
                    var right = IfBorder((int)cell.Borders.Item[Index.xlEdgeRight].LineStyle);
                    var bottom = IfBorder((int)cell.Borders.Item[Index.xlEdgeBottom].LineStyle);
                    var top = IfBorder((int)cell.Borders.Item[Index.xlEdgeTop].LineStyle);

                    HorizontalBorders[i, j] = top;
                    HorizontalBorders[i + 1, j] = bottom;
                    VerticalBorders[i, j] = left;
                    VerticalBorders[i, j + 1] = right;
                }
            }
        }
    }
}
