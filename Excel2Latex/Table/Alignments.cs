using static Excel2Latex.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Latex.Table
{
    public enum AlignmentFlag
    {
        Left,
        Right,
        Center
    }
    internal class Alignments
    {
        public AlignmentFlag[,] AlignmentFlags { get;}
        public Alignments(Excel.Range range)
        {
            var m = range.Rows.Count;
            var n = range.Columns.Count;

            AlignmentFlags = new AlignmentFlag[m, n];
            for (var i = 0; i < m; i++)
            {
                for (var j = 0; j < n; j++)
                {
                    var cell = range.Item[i + 1, j + 1];
                    var alignment = JudgeAlignment((int)cell.HorizontalAlignment);
                    AlignmentFlags[i, j] = alignment;
                }
            }
        }
    }
}
