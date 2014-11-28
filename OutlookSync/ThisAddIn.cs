using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NLog;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookSync
{
    public partial class ThisAddIn
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();

        public Dictionary<string, StoredAppointment> Appointments { get; set; }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new SyncRibbon();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {   
            var folder_path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookSync");
            if (!Directory.Exists(folder_path))
                Directory.CreateDirectory(folder_path);

            Appointments = new Dictionary<string, StoredAppointment>();
            var win = new MainWindow();
            win.Show();
            //SyncEngine.Execute();

            //var calendar = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            //logger.Trace(calendar.Name);
            //logger.Trace("Subfolders " + calendar.Folders.Count);
            //logger.Trace("Items " + calendar.Items.Count);

            //var items = GetItems(calendar.Items);
            //logger.Trace("Items from now and 1 year ahead: " + items.Count);
            //foreach (var item in items)
            //{
            //    logger.Trace("[{0}], {1} - {2} in {3} [{4}]", item.Subject, item.Start, item.End, item.Location);
            //}


            //calendar.Items.ItemAdd += ItemsOnItemAdd;
            //calendar.Items.ItemChange += ItemsOnItemChange;
            //calendar.Items.ItemRemove += ItemsOnItemRemove;

            //foreach (var item in calendar.Items)
            //{
            //    var appointment = item as Outlook.AppointmentItem;

            //    appointment.LastModificationTime

            //    if (appointment != null)
            //        logger.Trace("[{0}], {1} - {2} in {3}", appointment.Subject, appointment.Start, appointment.End, appointment.Location);
            //    else
            //        logger.Warn("Not an appointment");
            //}


            //var text = JsonConvert.SerializeObject(calendar.Items);
            //var base_dir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            //var file = Path.Combine(base_dir, "appointments.json");
            //File.WriteAllText(file, text);
        }

        //private static List<Outlook.AppointmentItem> GetItems(Outlook.Items items)
        //{
        //    var filter = "[Start] >= '" + DateTime.Now.ToString("g") + "' and [Start] <= '" + DateTime.Now.AddYears(1).ToString("g") + "'";

        //    items.IncludeRecurrences = true;
        //    items.Sort("[Start]", Type.Missing);

        //    return items.Restrict(filter)
        //                .Cast<object>()
        //                .OfType<Outlook.AppointmentItem>()
        //                .ToList();
        //}

        //private void ItemsOnItemRemove()
        //{
        //}

        //private void ItemsOnItemChange(object Item)
        //{
        //    throw new NotImplementedException();
        //}

        //private void ItemsOnItemAdd(object item)
        //{
        //    logger.Trace("item added");
        //}

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
