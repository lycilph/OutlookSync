using System;
using System.IO;
using System.Threading.Tasks;
using NLog;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSync
{
    public partial class ThisAddIn
    {
        private static readonly Logger log = LogManager.GetCurrentClassLogger();
        public static readonly string ApplicationName = "OutlookSync";

        // These are needed to keep the reference to the event handlers "live"
        private Outlook.MAPIFolder calendar_folder;
        private Outlook.Items calendar_items;

        public AddinSettings Settings { get; private set; }
        public SyncEngine SyncEngine { get; private set; }
        public SyncScheduler SyncScheduler { get; private set; }

        private async void ThisAddIn_Startup(object sender, EventArgs e)
        {
            log.Trace("Startup");

            var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ApplicationName);
            Settings = AddinSettings.Load(dir);
            SyncEngine = new SyncEngine(dir);
            SyncScheduler = new SyncScheduler();

            log.Trace("Attaching event handlers");

            calendar_folder = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            calendar_items = calendar_folder.Items;
            calendar_items.ItemAdd += item => RequestSync("Item added");
            calendar_items.ItemChange += item => RequestSync("Item changed");
            calendar_items.ItemRemove += () => RequestSync("Item removed");

            if (Settings.IsLoggedIn)
                await Task.Factory.StartNew(() => SyncEngine.Initialize());

            RequestSync("Addin started");
        }

        private void RequestSync(string message)
        {
            log.Trace(message);
            SyncScheduler.RequestSync();
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
