using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
