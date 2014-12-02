using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Requests;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSync
{
    public class SyncEngine
    {
        private const string ClientSecrets = "client_secrets.json";

        private readonly TaskCompletionSource<bool> tcs = new TaskCompletionSource<bool>();
        private readonly string dir;
        private CalendarService service;

        public Task IsReady { get { return tcs.Task; } }

        public SyncEngine(string dir)
        {
            this.dir = dir;
        }

        public void Initialize()
        {
            var secrets_file = Path.Combine(dir, ClientSecrets);
            UserCredential credential;
            using (var stream = new FileStream(secrets_file, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new[] { CalendarService.Scope.Calendar },
                    "user", CancellationToken.None,
                    new FileDataStore(dir, true)).Result;
            }

            service = new CalendarService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = ThisAddIn.ApplicationName,
            });
            
            tcs.SetResult(true);
        }

        public IEnumerable<GoogleCalendar> GetCalendars()
        {
            IsReady.Wait();

            var calendars = service.CalendarList.List().Execute();
            return calendars.Items.Select(i => new GoogleCalendar(i.Summary, i.Id));
        }

        public List<StoredAppointment> GetGoogleItems(string id, DateTime start, DateTime end)
        {
            var request = service.Events.List(id);
            request.TimeMin = start;
            request.TimeMax = end;
            return request.Execute().Items
                          .Select(e => new StoredAppointment(e))
                          .OrderBy(e => e.Start)
                          .ToList();
        }

        public List<StoredAppointment> GetOutlookItems(DateTime start, DateTime end)
        {
            var calendar = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            var items = calendar.Items;

            items.IncludeRecurrences = true;
            items.Sort("[Start]", Type.Missing);

            var filter = "[Start] >= '" + start.ToString("g") + "' and [Start] < '" + end.ToString("g") + "'";
            return items.Restrict(filter)
                        .Cast<object>()
                        .OfType<Outlook.AppointmentItem>()
                        .Select(a => new StoredAppointment(a))
                        .ToList();
        }

        public void AddGoogleItems(string id, IEnumerable<StoredAppointment> items)
        {
            var chunks = items.Chunk(50).ToList();
            foreach (var chunk in chunks)
            {
                var br = new BatchRequest(service);

                foreach (var appointment in chunk)
                {
                    var google_event = appointment.ToGoogleEvent();
                    var request = service.Events.Insert(google_event, id);
                    br.Queue<Event>(request, (r, e, i, m) =>
                    {
                        if (!m.IsSuccessStatusCode)
                            MessageBox.Show("Error: " + e.Message);
                    });
                }

                br.ExecuteAsync().Wait();
                Thread.Sleep(250);
            }
        }

        public void RemoveGoogleItems(string id, IEnumerable<StoredAppointment> items)
        {
            var chunks = items.Chunk(50).ToList();
            foreach (var chunk in chunks)
            {
                var br = new BatchRequest(service);

                foreach (var appointment in chunk)
                {
                    var request = service.Events.Delete(id, appointment.Id);
                    br.Queue<Event>(request, (r, e, i, m) =>
                    {
                        if (!m.IsSuccessStatusCode)
                            MessageBox.Show("Error: " + e.Message);
                    });
                }

                br.ExecuteAsync().Wait();
                Thread.Sleep(250);
            }
        }
    }
}
