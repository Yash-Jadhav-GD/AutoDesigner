using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

using Microsoft.Office.Tools;

namespace Autoformat_v1
{
    public partial class ThisAddIn
    {
        private CustomTaskPane myCustomTaskPane;
        public CustomTaskPane TaskPane => myCustomTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var myUserControl = new DesignSelectionTaskPane();
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl, "Design Tools");
            myCustomTaskPane.Width = 350;
        }

        public void HideTaskPane()
        {
            if (myCustomTaskPane != null)
            {
                try { myCustomTaskPane.Visible = false; } catch { }
            }
        }

        public void ToggleTaskPane()
        {
            if (myCustomTaskPane != null)
            {
                try
                {
                    myCustomTaskPane.Visible = !myCustomTaskPane.Visible;
                    return;
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // The taskpane was closed/deleted by the user, mark as null so it gets recreated
                    myCustomTaskPane = null;
                }
            }

            // If it was never created, or if it was destroyed, recreate it
            if (myCustomTaskPane == null)
            {
                var myUserControl = new DesignSelectionTaskPane();
                myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl, "Design Tools");
                myCustomTaskPane.Width = 350;
                myCustomTaskPane.Visible = true;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new FormattingRibbon();
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
