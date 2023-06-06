using System;
using System.Collections.Generic;
using System.Linq;
using Excel2Latex.Extensions;
using Excel = Microsoft.Office.Interop.Excel;
using Index = Microsoft.Office.Interop.Excel.XlBordersIndex;

namespace Excel2Latex.Table
{
    internal sealed class Table
    {
        public ActualAlignment[,] Alignments { get; }
        public bool[,] HorizontalBorders { get; }
        public bool[,] VerticalBorders { get; }
        public TextContext[,] TextContexts { get; }
        public bool[,] MergeAreas { get; }
        public int RowCount { get; }
        public int ColumnCount { get; }
        public ActualAlignment[] HeadAlignments { get; }
        public bool[] HeadBorders { get; }
        public Table(Excel.Range range)
        {
            #region 初始化属性

            RowCount = range.Rows.Count;
            ColumnCount = range.Columns.Count;

            Alignments = new ActualAlignment[RowCount, ColumnCount];
            HorizontalBorders = new bool[RowCount + 1, ColumnCount];
            VerticalBorders = new bool[RowCount, ColumnCount + 1];
            TextContexts = new TextContext[RowCount, ColumnCount];
            MergeAreas = new bool[RowCount, ColumnCount];
            HeadAlignments = new ActualAlignment[ColumnCount];
            HeadBorders = new bool[ColumnCount + 1];

            #endregion

            for (var i = 0; i < RowCount; i++)
            {
                for (var j = 0; j < ColumnCount; j++)
                {
                    var cell = range.Item[i + 1, j + 1];

                    #region 设置文字内容

                    TextContexts[i, j] = new TextContext(cell);

                    #endregion

                    #region 设置单元格对齐方式

                    Alignments[i, j] = TextContexts[i, j].Alignment;

                    #endregion

                    #region 设置单元格边框

                    HorizontalBorders[i, j] = ((int)cell.Borders.Item[Index.xlEdgeTop].LineStyle).IfBorder();//上
                    HorizontalBorders[i + 1, j] = ((int)cell.Borders.Item[Index.xlEdgeBottom].LineStyle).IfBorder();//下
                    VerticalBorders[i, j] = ((int)cell.Borders.Item[Index.xlEdgeLeft].LineStyle).IfBorder();//左
                    VerticalBorders[i, j + 1] = ((int)cell.Borders.Item[Index.xlEdgeRight].LineStyle).IfBorder();//右

                    #endregion

                    #region 设置合并区域

                    if ((bool)cell.MergeCells)
                    {
                        MergeAreas[i, j] = true;
                    }

                    #endregion
                }
            }

            SetHeadBorder(0);
            for (var i = 0; i < ColumnCount; i++)
            {
                SetHeadAlignment(i);//设置表头总的对齐方式
                SetHeadBorder(i + 1);//设置表头总的竖直边框
            }
        }
        public IEnumerable<Tuple<int, int>> GetContinuousHorizontalBorder(int rowNumber)
        {
            var start = -1;

            for (var i = 0; i < ColumnCount; i++)
            {
                var border = HorizontalBorders[rowNumber, i];

                if (!border) continue;
                if (start == -1)
                {
                    start = i;
                }

                if (i + 1 != ColumnCount && HorizontalBorders[rowNumber, i + 1]) continue;
                yield return new Tuple<int, int>(start + 1, i + 1);
                start = -1;
            }
        }
        private void SetHeadAlignment(int columnNumber)
        {
            var alignments = new ActualAlignment[RowCount];
            for (var i = 0; i < RowCount; i++)
            {
                alignments[i] = Alignments[i, columnNumber];
            }//提取某一列

            HeadAlignments[columnNumber] = alignments.GroupBy(alignment => alignment)
                .OrderByDescending(group => group.Count()).First().Key;
        }
        private void SetHeadBorder(int columnNumber)
        {
            var result = 0;
            for (var i = 0; i < RowCount; i++)
            {
                result += Convert.ToInt32(VerticalBorders[i, columnNumber]);
            }

            HeadBorders[columnNumber] = result > RowCount / 2;
        }
        public Tuple<int, int> GetCellMergeArea(int rowNumber, int columnNumber)
        {//TODO：还没有验证该算法的正确性
            var merge = MergeAreas[rowNumber, columnNumber];
            if (!merge)
            {
                return new Tuple<int, int>(0, 0);
            }

            if ((rowNumber > 0 && MergeAreas[rowNumber - 1, columnNumber]) || (columnNumber > 0 && MergeAreas[rowNumber, columnNumber - 1]))
            {
                return new Tuple<int, int>(0, 0);
            }

            var x = ColumnCount - columnNumber;
            var y = RowCount - rowNumber;
            var upper = x;
            for (var i = 1; i < upper; i++)
            {
                merge = MergeAreas[rowNumber, columnNumber + i];
                if (merge) continue;
                x = i;
                break;
            }

            upper = y;
            for (var i = 1; i < upper; i++)
            {
                merge = MergeAreas[rowNumber + i, columnNumber];
                if (merge) continue;
                y = i;
                break;
            }

            return new Tuple<int, int>(x, y);
        }
    }
}