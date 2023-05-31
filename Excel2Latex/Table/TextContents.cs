using System;
using System.Collections.Generic;
using System.Linq;
using static Excel2Latex.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Latex.Table
{
    public struct TextContent
    {
        private static readonly CommandSequenceExpression Expression;
        static TextContent()
        {
            var bold = new BoldExpression();
            var italic = new ItalicExpression();
            var underline = new UnderlineExpression();
            var translate = new TranslateExpression();
            var textColor = new TextColorExpression();
            Expression = new CommandSequenceExpression(textColor, bold, italic, underline, translate);
        }
        public string Text;
        public bool Bold;
        public bool Italic;
        public bool Underline;
        public Tuple<int, int, int> TextColor;
        public override string ToString()
        {
            return Expression.Interpret(this);
        }
    }
    internal class TextContents
    {
        public TextContent[,] Contents { get; }
        public TextContents(Excel.Range range)
        {
            var m = range.Rows.Count;
            var n = range.Columns.Count;

            Contents = new TextContent[m, n];
            for (var i = 0; i < m; i++)
            {
                for (var j = 0; j < n; j++)
                {
                    var cell = range.Item[i + 1, j + 1];
                    Contents[i, j] = new TextContent
                    {
                        Text = cell.Text,
                        Bold = cell.Font.Bold,
                        Italic = cell.Font.Italic,
                        Underline = IfUnderline(cell.Font.Underline),
                        TextColor = Int2Rgb((int)cell.Font.Color)
                    };
                }
            }
        }
    }
    internal abstract class AbstractExpression
    {
        public abstract string Interpret(TextContent context);
    }
    internal class TextColorExpression : AbstractExpression
    {
        private static readonly Tuple<int, int, int> DefaultColor = new Tuple<int, int, int>(0, 0, 0);
        public override string Interpret(TextContent context)
        {
            var color = context.TextColor;
            if (Equals(color, DefaultColor))
            {
                return "";
            }
            var (r, g, b) = color;
            return $@"\textcolor[rgb]{{{r},{g},{b}}}";
        }
    }
    internal class UnderlineExpression : AbstractExpression
    {
        public override string Interpret(TextContent context)
        {
            return context.Underline ? @"\underline" : "";
        }
    }
    internal class ItalicExpression : AbstractExpression
    {
        public override string Interpret(TextContent context)
        {
            return context.Italic ? @"\textit" : "";
        }
    }
    internal class BoldExpression : AbstractExpression
    {
        public override string Interpret(TextContent context)
        {
            return context.Bold ? @"\bold" : "";
        }
    }
    internal class CommandSequenceExpression : AbstractExpression
    {
        private const string Separator = "{";
        public List<AbstractExpression> Expressions { get; set; }
        public CommandSequenceExpression(params AbstractExpression[] expressions)
        {
            Expressions = new List<AbstractExpression>();
            foreach (var expression in expressions)
            {
                Expressions.Add(expression);
            }
        }
        public override string Interpret(TextContent context)
        {
            var resultList = Expressions.Select(expression => expression.Interpret(context)).Where(temp => temp != "").ToList();//返回空字符串则说明命令无需设置
            return string.Join(Separator, resultList) + new string('}', resultList.Count - 1);
        }
    }
    internal class TranslateExpression : AbstractExpression
    {
        private static readonly Dictionary<string, string> EscapeDictionary = new Dictionary<string, string>
        {
            [@"\"] = @"\textbackslash{}",
            ["$"] = @"\$",
            ["^"] = @"\^",
            ["_"] = @"\_"
        };
        public static bool TranslateRequired = true;
        public override string Interpret(TextContent context)
        {
            var text = context.Text;
            if (TranslateRequired)
            {
                text = EscapeDictionary.Aggregate(text,
                    (current, value) => current.Replace(value.Key, value.Value));//不确定顺序是否会产生影响
            }

            text = text.Replace("%", @"\%");
            if (text.Contains("\n"))
            {
                return @"\makecell{" + text.Replace("\n", @"\\") + "}";
            }

            return text;
        }
    }
}
