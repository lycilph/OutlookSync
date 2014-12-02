using System;
using System.IO;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;

namespace OutlookSync
{
    public partial class ThisAddIn
    {
        public static readonly string ApplicationName = "OutlookSync";

        public AddinSettings Settings { get; private set; }
        public SyncEngine SyncEngine { get; private set; }

        private async void ThisAddIn_Startup(object sender, EventArgs e)
        {
            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ApplicationName);
            Settings = AddinSettings.Load(dir);
            SyncEngine = new SyncEngine(dir);

            if (Settings.IsInitialized)
                await Task.Factory.StartNew(() => SyncEngine.Initialize());
        }

        // This is NOT called at all! (see http://msdn.microsoft.com/en-us/library/office/ee720183.aspx#OL2010AdditionalShutdownChanges_BestPracticesforAddinShutdownforDevelopers)
        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SyncRibbon();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
