using System.Linq;
using System.Text;

namespace Excel2Latex.Builder
{
    internal sealed class StandardTableBuilder : LatexTableBuilder
    {
        public StandardTableBuilder(Microsoft.Office.Interop.Excel.Range range) : base(range)
        {
        }
        public override void StartTableEnvironment()
        {
            Builder.AppendLine(@"\begin{table}[htbp]");
        }
        public override void EndTableEnvironment()
        {
            Builder.Append(@"\end{table}");
        }
        public override void BuildHorizontalLine(int rowNumber)
        {
            var builder = new StringBuilder();
            var result = Table.GetContinuousHorizontalBorder(rowNumber).ToList();
            var length = result.Count;

            switch (length)
            {
                case 0:
                    break;
                case 1:
                    var single = result.First();
                    if (single.Item1 == 1 && single.Item2 == Table.ColumnCount)
                    {
                        builder.Append(@"\hline");
                    }
                    else
                    {
                        builder.Append($@"\cline[{single.Item1}-{single.Item2}]");
                    }
                    break;
                default:
                    foreach (var tuple in result)
                    {
                        builder.Append($@"\cline[{tuple.Item1}-{tuple.Item2}]");
                    }
                    break;
            }

            builder.Insert(0, "\t\t");
            builder.Append("\n");
            Builder.Append(builder);
        }
        public override void EndTabularEnvironment()
        {
            Builder.AppendLine("\t\\end{tabular}");
        }
        public override void StartTabularEnvironment()
        {
            var builder = new StringBuilder("\t\\begin{tabular}{");

            string FuncBorder(bool border) => border ? "|" : "";

            builder.Append(FuncBorder(Table.HeadBorders[0]));

            for (var i = 0; i < Table.ColumnCount; i++)
            {
                builder.Append(Table.HeadAlignments[i].ToString().ToLower());
                builder.Append(FuncBorder(Table.HeadBorders[i + 1]));
            }

            builder.Append("}\n");
            Builder.Append(builder);
        }
        public override string GetResult()
        {
            StartTableEnvironment();
            StartTabularEnvironment();

            for (var i = 0; i < Table.RowCount + 1; i++)
            {
                BuildHorizontalLine(i);
            }

            EndTabularEnvironment();
            EndTableEnvironment();

            return Builder.ToString();
        }
    }
}