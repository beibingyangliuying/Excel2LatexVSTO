using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Latex.Builder
{
    internal abstract class LatexTableBuilder
    {
        protected StringBuilder Builder = new StringBuilder();
        protected Table.Table Table;
        protected LatexTableBuilder(Excel.Range range)
        {
            Table = new Table.Table(range);
        }
        public virtual void StartTableEnvironment() { }
        public virtual void StartTabularEnvironment() { }
        public virtual void BuildHorizontalLine(int rowNumber) { }
        public virtual void BuildRow(int rowNumber) { }
        public virtual void EndTabularEnvironment() { }
        public virtual void EndTableEnvironment() { }
        public virtual string GetResult()
        {
            return "";
        }
    }
}