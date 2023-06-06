using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Latex.Builder
{
    internal abstract class LatexTableBuilder
    {
        protected StringBuilder Builder = new StringBuilder();
        protected Table.Table Table;
        protected string Caption;
        protected string Label;
        protected LatexTableBuilder(Excel.Range range, string caption, string label)
        {
            Table = new Table.Table(range);
            Caption = caption;
            Label = label;
        }
        public virtual void StartTableEnvironment() { }
        public virtual void StartTabularEnvironment() { }
        public virtual void BuildHorizontalLine(int rowNumber) { }
        public virtual void BuildRow(int rowNumber) { }
        public virtual void BuildCaption() { }
        public virtual void BuildLabel() { }
        public virtual void EndTabularEnvironment() { }
        public virtual void EndTableEnvironment() { }
        public virtual string GetResult()
        {
            return Builder.ToString();
        }
    }
}