using System;
using System.Collections.Generic;
using System.Linq;

namespace Excel2Latex
{
    internal abstract class AbstractExpression
    {
        public abstract string InterpretRangeContext(CellContext context);
    }

    internal class TextColorExpression : AbstractExpression
    {
        private static readonly Tuple<int, int, int> DefaultColor = new Tuple<int, int, int>(0, 0, 0);
        public override string InterpretRangeContext(CellContext context)
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

    /// <summary>
    /// 常量字符串
    /// </summary>
    internal class LiteralExpression : AbstractExpression
    {
        private readonly string _literal;
        public LiteralExpression(string literal)
        {
            _literal = literal;
        }
        public override string InterpretRangeContext(CellContext context)
        {
            return _literal;
        }
    }

    /// <summary>
    /// 判断单元格是否有下划线
    /// </summary>
    internal class UnderlineExpression : AbstractExpression
    {
        public override string InterpretRangeContext(CellContext context)
        {
            return context.Underline ? @"\underline" : "";
        }
    }

    /// <summary>
    /// 判断单元格是否斜体
    /// </summary>
    internal class ItalicExpression : AbstractExpression
    {
        public override string InterpretRangeContext(CellContext context)
        {
            return context.Italic ? @"\textit" : "";
        }
    }

    /// <summary>
    /// 判断单元格字体是否加粗
    /// </summary>
    internal class BoldExpression : AbstractExpression
    {
        public override string InterpretRangeContext(CellContext context)
        {
            return context.Bold ? @"\bold" : "";
        }
    }

    /// <summary>
    /// 合成命令序列
    /// </summary>
    internal class CommandSequenceExpression : AbstractExpression
    {
        private const string Separator = "{";
        public List<AbstractExpression> Expressions { get; set; }
        public CommandSequenceExpression(List<AbstractExpression> expressions)
        {
            Expressions = expressions;
        }
        public override string InterpretRangeContext(CellContext context)
        {
            var resultList = Expressions.Select(expression => expression.InterpretRangeContext(context)).Where(temp => temp != "").ToList();//返回空字符串则说明命令无需设置
            return string.Join(Separator, resultList) + new string('}', resultList.Count - 1);
        }
    }

    /// <summary>
    /// 转义单元格内容
    /// </summary>
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
        public override string InterpretRangeContext(CellContext context)
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
                return @"\makecell{" + text.Replace("\n",@"\\") + "}";
            }

            return text;
        }
    }
}
