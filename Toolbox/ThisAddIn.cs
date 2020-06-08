using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace Toolbox
{
    public partial class ThisAddIn
    {
        private readonly Ribbon1 _ribbon = new Ribbon1();
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return _ribbon;
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.Application.WindowSelectionChange += Application_WindowSelectionChange;
        }
        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            //_ribbon.InvalidateToggle();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
