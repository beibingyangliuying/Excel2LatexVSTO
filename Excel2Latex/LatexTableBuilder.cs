using System.Text;

namespace Excel2Latex
{
    internal class LatexDirector
    {
        private LatexTableBuilder _tableBuilder;
        public string Construct()
        {
            return _tableBuilder.GetResult();
        }
    }
    internal abstract class LatexTableBuilder
    {
        protected StringBuilder Builder;
        public virtual void StartTableEnvironment() { }
        public virtual void StartTabularEnvironment() { }
        public virtual void BuildHorizontalLine() { }
        public virtual void BuildRow() { }
        public virtual void EndTabularEnvironment() { }
        public virtual void EndTableEnvironment() { }
        public virtual string GetResult()
        {
            return "";
        }
    }
    internal class StandardTableBuilder : LatexTableBuilder
    {
        public override void StartTableEnvironment()
        {
            Builder.AppendLine(@"\begin{table}[htbp]");
        }
        public override void EndTableEnvironment()
        {
            Builder.Append(@"\end{table}");
        }
    }
}