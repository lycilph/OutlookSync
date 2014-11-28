using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using NLog;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookSync
{
    public class SyncEngine
    {
        private static readonly Logger logger = LogManager.GetCurrentClassLogger();
        private const string application_name = "OutlookSync";

        private CalendarService service;

        public string BaseDir { get; private set; }

        public SyncEngine()
        {
            BaseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), application_name);
        }

        public void Initialize()
        {
            var secrets_file = Path.Combine(BaseDir, "client_secrets.json");

            UserCredential credential;
            using (var stream = new FileStream(secrets_file, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new[] { CalendarService.Scope.Calendar },
                    "user", CancellationToken.None,
                    new FileDataStore(BaseDir, true)).Result;
            }

            service = new CalendarService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = application_name,
            });
        }

        public List<StoredAppointment> GetOutlookItems(DateTime start, DateTime end)
        {
            var calendar = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            var items = calendar.Items;

            items.IncludeRecurrences = true;
            items.Sort("[Start]", Type.Missing);

            var filter = "[Start] >= '" + start.ToString("g") + "' and [Start] <= '" + end.ToString("g") + "'";
            return items.Restrict(filter)
                        .Cast<object>()
                        .OfType<Outlook.AppointmentItem>()
                        .Select(a => new StoredAppointment(a))
                        .ToList();
        }

        public List<StoredAppointment> GetGoogleItems(string id, DateTime start, DateTime end)
        {
            var request = service.Events.List(id);
            request.TimeMin = start;
            request.TimeMax = end;
            return request.ExecuteAsync().Result.Items
                          .Select(e => new StoredAppointment(e))
                          .ToList();
        }

        public IEnumerable<string> GetGoogleCalendars()
        {
            var request = service.CalendarList.List();
            var response = request.ExecuteAsync().Result;
            return response.Items.Select(i => i.Summary + ":" + i.Id);
        }
    }
}
