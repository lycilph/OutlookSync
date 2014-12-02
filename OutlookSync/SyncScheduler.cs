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

        public bool SyncRequested { get; private set; }

        public SyncScheduler()
        {
            sync_engine = Globals.ThisAddIn.SyncEngine;

            var interval = TimeSpan.FromMinutes(10).TotalMilliseconds;
            set initial interval to 5 sec, then set correct interval in then Sync method
            var timer = new Timer(interval);
            timer.Elapsed += Sync;
            timer.Start();
        }

        private async void Sync(object sender, ElapsedEventArgs args)
        {
            log.Trace("Checking sync requests");
             interval = TimeSpan.FromMinutes(10).TotalMilliseconds;
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
    }
}
