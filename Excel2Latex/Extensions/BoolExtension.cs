namespace Excel2Latex.Extensions
{
    internal static class BoolExtension
    {
        public static string ToString(this bool value)
        {
            return value ? "|" : "";
        }
    }
}
