namespace Excel2Latex.Builder
{
    internal sealed class LatexDirector
    {
        public LatexTableBuilder TableBuilder { get; set; }
        public LatexDirector(LatexTableBuilder tableBuilder)
        {
            TableBuilder = tableBuilder;
        }
        public string Construct()
        {
            return TableBuilder.GetResult();
        }
    }
}