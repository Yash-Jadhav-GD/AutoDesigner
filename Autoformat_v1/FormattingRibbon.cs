using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace Autoformat_v1
{
    [ComVisible(true)]
    public class FormattingRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public FormattingRibbon()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Autoformat_v1.FormattingRibbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnOpenPanelClicked(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ToggleTaskPane();
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }
    }
}
