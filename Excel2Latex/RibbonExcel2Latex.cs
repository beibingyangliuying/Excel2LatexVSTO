using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel2Latex.Builder;

namespace Excel2Latex
{
    public partial class RibbonExcel2Latex
    {
        private void RibbonExcel2Latex_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonTransfer_Click(object sender, RibbonControlEventArgs e)
        {
            var range = Globals.ThisAddIn.Application.Selection;
            var tableBuilder = new StandardTableBuilder(range);
            var director = new LatexDirector(tableBuilder);
            MessageBox.Show(director.Construct());
        }
    }
}
