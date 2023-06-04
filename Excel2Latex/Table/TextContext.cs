using System;
using System.Collections.Generic;
using System.Linq;
using Excel2Latex.Extensions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2Latex.Table
{
    internal readonly struct TextContext
    {
        private static readonly CommandSequenceExpression Expression;
        static TextContext()
        {
            var bold = new BoldExpression();
            var italic = new ItalicExpression();
            var underline = new UnderlineExpression();
            var translate = new TranslateExpression();
            var textColor = new TextColorExpression();
            Expression = new CommandSequenceExpression(textColor, bold, italic, underline, translate);
        }
        private readonly ExcelAlignment _alignment;
        public string Text { get; }
        public bool Bold { get; }
        public bool Italic { get; }
        public bool Underline { get; }
        public Tuple<int, int, int> TextColor { get; }
        public TextContext(Excel.Range range)
        {
            Text = range.Text.Trim();
            Bold = range.Font.Bold;
            Italic = range.Font.Italic;
            Underline = ((int)range.Font.Underline).IfUnderline();
            TextColor = ((int)range.Font.Color).ToRgb();
            _alignment = (ExcelAlignment)(int)range.HorizontalAlignment;
        }
        public override string ToString()
        {
            return Text == "" ? "" : Expression.Interpret(this);
        }
        public ActualAlignment Alignment
        {
            get
            {
                switch (_alignment)
                {
                    case ExcelAlignment.Center:
                        return ActualAlignment.C;
                    case ExcelAlignment.Left:
                        return ActualAlignment.L;
                    case ExcelAlignment.Right:
                        return ActualAlignment.R;
                    case ExcelAlignment.General:
                        if (double.TryParse(Text, out _) || DateTime.TryParse(Text, out _))
                        {
                            return ActualAlignment.R;
                        }
                        return ActualAlignment.L;
                    default:
                        return ActualAlignment.C;
                }
            }
        }
    }
    internal abstract class AbstractExpression
    {
        public abstract string Interpret(TextContext context);
    }
    internal sealed class TextColorExpression : AbstractExpression
    {
        private static readonly Tuple<int, int, int> DefaultColor = new Tuple<int, int, int>(0, 0, 0);
        public override string Interpret(TextContext context)
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
    internal sealed class UnderlineExpression : AbstractExpression
    {
        public override string Interpret(TextContext context)
        {
            return context.Underline ? @"\underline" : "";
        }
    }
    internal sealed class ItalicExpression : AbstractExpression
    {
        public override string Interpret(TextContext context)
        {
            return context.Italic ? @"\textit" : "";
        }
    }
    internal sealed class BoldExpression : AbstractExpression
    {
        public override string Interpret(TextContext context)
        {
            return context.Bold ? @"\bold" : "";
        }
    }
    internal sealed class CommandSequenceExpression : AbstractExpression
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
        public override string Interpret(TextContext context)
        {
            var resultList = Expressions.Select(expression => expression.Interpret(context)).Where(temp => temp != "").ToList();//返回空字符串则说明命令无需设置
            return string.Join(Separator, resultList) + new string('}', resultList.Count - 1);
        }
    }
    internal sealed class TranslateExpression : AbstractExpression
    {
        private static readonly Dictionary<string, string> EscapeDictionary = new Dictionary<string, string>
        {
            [@"\"] = @"\textbackslash{}",
            ["$"] = @"\$",
            ["^"] = @"\^",
            ["_"] = @"\_",
            ["#"] = @"\#"
        };
        public static bool TranslateRequired = true;
        public override string Interpret(TextContext context)
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
