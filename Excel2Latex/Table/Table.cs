using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Latex.Table
{
    internal class Table
    {
        public Alignments Alignments { get; }
        public Borders Borders { get; }
        public TextContents Contents { get; }
        public Table(Excel.Range range)
        {
            Alignments = new Alignments(range);
            Borders = new Borders(range);
            Contents = new TextContents(range);
        }
    }
}