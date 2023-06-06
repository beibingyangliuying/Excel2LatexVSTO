using System;

namespace Excel2Latex
{
    public partial class ThisAddIn
    {
        private static void ThisAddIn_Startup(object sender, EventArgs e)
        {
        }

        private static void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
