using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel2Latex
{
    internal class BordersContext
    {
    }

    internal class CellContext:AbstractContext
    {
        public CellContext(Microsoft.Office.Interop.Excel.Range range)
        {
            Range = range;
        }

        public Microsoft.Office.Interop.Excel.Range Range { get; }
        public string Text => Range.Text;
        public bool Bold => Range.Font.Bold;
        public bool Italic => Range.Font.Italic;
        public bool Underline => Range.Font.Underline != -4142; //TODO：-4142是无下划线时的默认值，将来可能更改
        public Tuple<int, int, int> TextColor => Utilities.Int2Rgb((int)Range.Font.Color);
        public override object Accept(AbstractExpression expression)
        {
            return expression.InterpretRangeContext(this);
        }
    }

    internal abstract class AbstractContext
    {
        public abstract object Accept(AbstractExpression expression);
    }
}