using System;

namespace Excel2Latex.Extensions
{
    internal static class IntExtension
    {
        public static Tuple<int, int, int> ToRgb(this int number)
        {
            var r = number % 256;
            var g = number / 256 % 256;
            var b = number / (256 * 256) % 256;
            return new Tuple<int, int, int>(r, g, b);
        }
        public static bool IfUnderline(this int number)
        {
            return number != -4142;
        }
        public static bool IfBorder(this int number)
        {
            return number != -4142;
        }
    }
}
