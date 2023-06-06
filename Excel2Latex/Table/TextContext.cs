using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
            return IsEmpty() ? "" : Expression.Interpret(this);
        }
        public bool IsEmpty()
        {
            return string.IsNullOrEmpty(Text);
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
                        {//数值型或日期型的数据默认居右对齐
                            return ActualAlignment.R;
                        }
                        return ActualAlignment.L;//文本型数据默认居左对齐
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
        private AbstractExpression[] Expressions { get; }
        public CommandSequenceExpression(params AbstractExpression[] expressions)
        {
            Expressions = expressions;
        }
        public override string Interpret(TextContext context)
        {
            var values = Expressions.Select(e => e.Interpret(context)).Where(s => s != "").ToList();//返回空字符串则说明命令无需设置
            return string.Join("{", values) + new string('}', values.Count - 1);
        }
    }
    internal sealed class TranslateExpression : AbstractExpression
    {
        private static readonly Dictionary<string, string> NecessaryEscapeDictionary = new Dictionary<string, string>
        {
            ["#"] = @"\#",
            ["%"] = @"\%"
        };
        private static readonly Dictionary<string, string> UnnecessaryEscapeDictionary = new Dictionary<string, string>
        {
            [@"\"] = @"\textbackslash{}",
            ["$"] = @"\$",
            ["^"] = @"\^",
            ["_"] = @"\_"
        };
        public static bool TranslateRequired = true;
        public override string Interpret(TextContext context)
        {
            var builder = new StringBuilder(context.Text);
            if (TranslateRequired)
            {//必须首先转换\字符
                builder = UnnecessaryEscapeDictionary.Aggregate(builder,
                    (b, pair) => b.Replace(pair.Key, pair.Value));
            }

            builder = NecessaryEscapeDictionary.Aggregate(builder, (b, pair) => b.Replace(pair.Key, pair.Value));

            if (!context.Text.Contains("\n")) return builder.ToString();

            builder.Replace("\n", @"\\");
            builder.Insert(0, @"\makecell{");
            builder.Append("}");
            return builder.ToString();
        }
    }
}
