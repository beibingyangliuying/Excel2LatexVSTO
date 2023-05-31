using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Latex
{
    internal class LatexContext
    {
        private string[,] _cells;
        private string[] _horizontalLines;
        private readonly CommandSequenceExpression _cellTextExpression;
        private readonly CellContext _cellContext;

        public LatexContext(CellContext cellContext)
        {
            _cellContext = cellContext;
            _cells = new string[_cellContext.Range.Rows.Count, _cellContext.Range.Columns.Count];
            _horizontalLines = new string[_cellContext.Range.Rows.Count + 1];

            var bold = new BoldExpression();
            var italic = new ItalicExpression();
            var underline = new UnderlineExpression();
            var translate = new TranslateExpression();
            var textColor = new TextColorExpression();

            _cellTextExpression = new CommandSequenceExpression(new List<AbstractExpression>
            {
                textColor,
                bold,
                italic,
                underline,
                translate
            });

        }

        public string GenerateLatexCode()
        {
            var builder = new StringBuilder();
            for (var i = 0; i < _cells.GetLength(0); i++)
            {
                for (var j = 0; j < _cells.GetLength(1); j++)
                {
                    var temp = new CellContext(_cellContext.Range.Item[i + 1, j + 1]);//偏移量必须从1开始
                    _cells[i, j] = _cellTextExpression.InterpretRangeContext(temp);
                    builder.Append(_cells[i, j] + "&");
                }

                builder.Append("\n");
            }

            return builder.ToString();
        }
    }
}
