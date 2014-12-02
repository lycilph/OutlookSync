using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new SyncRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace OutlookSync
{
    [ComVisible(true)]
    public class SyncRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public SyncRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbon_id)
        {
            return GetResourceText("OutlookSync.SyncRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void SyncUICallback(Office.IRibbonControl control)
        {
            var win = new SyncWindow();
            win.Show();
        }

        public void SyncAllCallback(Office.IRibbonControl control)
        {
        }

        public void SettingsCallback(Office.IRibbonControl control)
        {
            var win = new SettingsWindow();
            win.Show();
        }

        public void Ribbon_Load(Office.IRibbonUI ribbon_ui)
        {
            ribbon = ribbon_ui;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string name_to_find)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resource_names = asm.GetManifestResourceNames();
            var name = resource_names.Single(n => string.Compare(name_to_find, n, StringComparison.OrdinalIgnoreCase) == 0);

            // ReSharper disable once AssignNullToNotNullAttribute
            using (var resource_reader = new StreamReader(asm.GetManifestResourceStream(name)))
            {
                return resource_reader.ReadToEnd();
            }
        }

        #endregion
    }
}
