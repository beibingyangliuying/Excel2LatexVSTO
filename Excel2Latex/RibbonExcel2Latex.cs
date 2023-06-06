using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel2Latex.Builder;

namespace Excel2Latex
{
    public partial class RibbonExcel2Latex
    {
        private void RibbonExcel2Latex_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ButtonTransfer_Click(object sender, RibbonControlEventArgs e)
        {
            var range = Globals.ThisAddIn.Application.Selection;
            var tableBuilder = new StandardTableBuilder(range,"caption","caption");
            var director = new LatexDirector(tableBuilder);

            var form = new FormLatexBuilder();
            form.SetResult(director.Construct());
            form.ShowDialog();
        }
    }
}
