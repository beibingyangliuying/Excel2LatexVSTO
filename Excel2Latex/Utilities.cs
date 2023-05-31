using System;
using Excel2Latex.Table;

namespace Excel2Latex
{
    internal static class Utilities
    {
        public static Tuple<int, int, int> Int2Rgb(int number)
        {
            var r = number % 256;
            var g = number / 256 % 256;
            var b = number / (256 * 256) % 256;
            return new Tuple<int, int, int>(r, g, b);
        }
        public static bool IfUnderline(int number)
        {
            return number != -4142;
        }
        public static bool IfBorder(int number)
        {
            return number != -4142;
        }
        public static AlignmentFlag JudgeAlignment(int number)
        {
            switch (number)
            {
                case -4108:
                    return AlignmentFlag.Center;
                case -4131:
                    return AlignmentFlag.Left;
                case -4152:
                    return AlignmentFlag.Right;
                default:
                    //throw new InvalidEnumArgumentException(nameof(number), number, typeof(AlignmentFlag));
                    Console.WriteLine(@"使用默认的居中对齐");
                    return AlignmentFlag.Center;
            }
        }
    }
}
