using System;
using System.Threading.Tasks;
using System.Timers;
using NLog;

namespace OutlookSync
{
    public class SyncScheduler
    {
        private static readonly Logger log = LogManager.GetCurrentClassLogger();
        private readonly SyncEngine sync_engine;
        private readonly AddinSettings settings;
        private readonly Timer timer;
        private bool first_time = true;

        public bool SyncRequested { get; private set; }

        public SyncScheduler()
        {
            sync_engine = Globals.ThisAddIn.SyncEngine;
            settings = Globals.ThisAddIn.Settings;

            var interval = TimeSpan.FromSeconds(30).TotalMilliseconds;
            timer = new Timer(interval);
            timer.Elapsed += Sync;
            timer.Start();
        }

        private async void Sync(object sender, ElapsedEventArgs args)
        {
            log.Trace("Checking sync requests");

            if (first_time)
            {
                log.Trace("Setting interval to " + settings.SchedulerInterval + " minutes");
                first_time = false;
                timer.Interval = TimeSpan.FromMinutes(settings.SchedulerInterval).TotalMilliseconds;
            }
             
            if (!SyncRequested) 
                return;

            log.Trace("Syncing started");

            SyncRequested = false;
            await Task.Factory.StartNew(() => sync_engine.Sync());

            log.Trace("Syncing done");
        }

        public void RequestSync()
        {
            if (!Globals.ThisAddIn.Settings.IsLoggedIn) 
                return;

            log.Trace("Sync requested");
            SyncRequested = true;
        }

        public void Start()
        {
            log.Trace("Starting");
            timer.Interval = TimeSpan.FromMinutes(settings.SchedulerInterval).TotalMilliseconds;
            timer.Start();
        }

        public void Stop()
        {
            log.Trace("Stopping");
            timer.Stop();
        }

        public void Restart()
        {
            Stop();
            Start();
        }
    }
}
