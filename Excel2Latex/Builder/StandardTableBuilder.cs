using System;
using System.Linq;
using System.Text;
using Excel2Latex.Extensions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Latex.Builder
{
    internal sealed class StandardTableBuilder : LatexTableBuilder
    {
        public StandardTableBuilder(Excel.Range range, string caption, string label) : base(range, caption, label)
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

            builder.Append(BoolExtension.ToString(Table.HeadBorders[0]));

            for (var i = 0; i < Table.ColumnCount; i++)
            {
                builder.Append(Table.HeadAlignments[i].ToString().ToLower());
                builder.Append(BoolExtension.ToString(Table.HeadBorders[i + 1]));
            }

            builder.Append("}\n");
            Builder.Append(builder);
        }
        public override void BuildRow(int rowNumber)
        {//TODO：需要重新设计
            var builder = new StringBuilder();
            var defaultMergeArea = new Tuple<int, int>(0, 0);

            for (var i = 0; i < Table.ColumnCount; i++)
            {
                var textContext = Table.TextContexts[rowNumber, i];
                //if (textContext.IsEmpty())
                //{
                //    builder.Append("&");
                //    continue;
                //}

                var text = new StringBuilder(textContext.ToString());

                var mergeArea = Table.GetCellMergeArea(rowNumber, i);
                var left = Table.VerticalBorders[rowNumber, i];
                var right = Table.VerticalBorders[rowNumber, i + 1];

                if (!Equals(mergeArea, defaultMergeArea))
                {
                    if (mergeArea.Item1 != 1)
                    {
                        text.Insert(0,
                            $@"\multicolumn{{{mergeArea.Item1}}}{{{BoolExtension.ToString(left)}{textContext.Alignment.ToString().ToLower()}{BoolExtension.ToString(right)}}}{{");
                        text.Append("}");
                    }

                    if (mergeArea.Item2 != 1)
                    {
                        text.Insert(0,
                            $@"\multirow{{{mergeArea.Item2}}}{{{BoolExtension.ToString(left)}{textContext.Alignment.ToString().ToLower()}{BoolExtension.ToString(right)}}}{{");
                        text.Append("}");
                    }
                }

                builder.Append(text);
                if (i + 1 != Table.ColumnCount)
                {
                    builder.Append("&");
                }
            }

            builder.AppendLine(@"\bigstrut \\");
            Builder.Append(builder);
        }
        public override void BuildCaption()
        {
            Builder.AppendLine($"\t\\caption{{{Caption}}}");
        }
        public override void BuildLabel()
        {
            Builder.AppendLine($"\t\\label{{tab:{Label}}}");
        }
        public override string GetResult()
        {
            StartTableEnvironment();
            StartTabularEnvironment();

            BuildCaption();
            BuildLabel();

            BuildHorizontalLine(0);
            for (var i = 0; i < Table.RowCount; i++)
            {
                BuildRow(i);
                BuildHorizontalLine(i + 1);
            }

            EndTabularEnvironment();
            EndTableEnvironment();

            return Builder.ToString();
        }
    }
}