using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel2Latex.Table;

namespace UnitTestProject
{
    [TestClass]
    public class UnitTestUtilities
    {
        [TestMethod]
        public void TestGetHorizontalBorder()
        {
            var borders = new[,]
            {
                {true,false,false,true,true,false,true,true},
                {false,false,false,false,false,false,false,false},
                {true,true,true,true,true,true,true,true},
                {true,false,true,false,true,false,true,false},
                {false,false,true,true,false,false,true,true}
            };
            var start = -1;
            var rowCount = borders.GetLength(0);
            var columnCount = borders.GetLength(1);
            var result = new List<Tuple<int, int>>();

            for (var i = 0; i < rowCount; i++)
            {
                for (var j = 0; j < columnCount; j++)
                {
                    var border = borders[i, j];

                    if (!border) continue;
                    if (start == -1)
                    {
                        start = j;
                    }

                    if (j + 1 != columnCount && borders[i, j+1]) continue;
                    result.Add(new Tuple<int, int>(start + 1, j + 1));
                    start = -1;
                }
            }

            foreach (var tuple in result)
            {
                Console.WriteLine(tuple);
            }
        }

        [TestMethod]
        public void TestGetVerticalBorder()
        {
            var borders = new[,]
            {
                {true,false,false,true,true,false,true,true},
                {false,false,false,false,false,false,false,false},
                {true,true,true,true,true,true,true,true},
                {true,false,true,false,true,false,true,false},
                {false,false,true,true,false,false,true,true}
            };
            var start = -1;
            var rowCount = borders.GetLength(0);
            var columnCount = borders.GetLength(1);
            var result = new List<Tuple<int, int>>();

            for (var i = 0; i < columnCount; i++)
            {
                for (var j = 0; j < rowCount; j++)
                {
                    var border = borders[j, i];

                    if (!border) continue;
                    if (start == -1)
                    {
                        start = j;
                    }

                    if (j + 1 != rowCount && borders[j + 1, i]) continue;
                    result.Add(new Tuple<int, int>(start + 1, j + 1));
                    start = -1;
                }
            }

            foreach (var tuple in result)
            {
                Console.WriteLine(tuple);
            }
        }

        //[TestMethod]
        //public void TestGetColumnHorizontalAlignment()
        //{
        //    var alignments = new ExcelAlignment[,]
        //    {
        //        {ExcelAlignment.Center,ExcelAlignment.Left,ExcelAlignment.Right,ExcelAlignment.Center,ExcelAlignment.Left},
        //        {ExcelAlignment.Center,ExcelAlignment.Right,ExcelAlignment.Center,ExcelAlignment.Center,ExcelAlignment.Right},
        //        {ExcelAlignment.Right,ExcelAlignment.Left,ExcelAlignment.Left,ExcelAlignment.Left,ExcelAlignment.Center},
        //        {ExcelAlignment.Left,ExcelAlignment.Right,ExcelAlignment.Right,ExcelAlignment.Center,ExcelAlignment.Left}
        //    };
        //    var rowCount = alignments.GetLength(0);
        //    var columnCount=alignments.GetLength(1);

        //    for (var i = 0; i < columnCount; i++)
        //    {
        //        var temp = new List<ExcelAlignment>();
        //        for (var j = 0; j < rowCount; j++)
        //        {
        //            temp.Add(alignments[j,i]);
        //        }

        //        var tempResult = temp.GroupBy(alignment => alignment).OrderByDescending(group=>group.Count()).ToList();
        //        foreach (var item in tempResult)
        //        {
        //            Console.WriteLine($"{item.Key}:{item.Count()}");
        //        }
        //        Console.WriteLine();
        //    }
        //}
    }
}
