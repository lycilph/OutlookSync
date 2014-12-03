using System.IO;
using Newtonsoft.Json;
using NLog;

namespace OutlookSync
{
    public class AddinSettings : ObservableObject
    {
        private static readonly Logger log = LogManager.GetCurrentClassLogger();
        private const string Filename = "settings.json";
        private string base_dir;

        private bool is_logged_in;
        public bool IsLoggedIn
        {
            get { return is_logged_in; }
            set
            {
                if (value.Equals(is_logged_in)) return;
                is_logged_in = value;
                OnPropertyChanged();
            }
        }

        private int sync_window;
        public int SyncWindow
        {
            get { return sync_window; }
            set
            {
                if (value.Equals(sync_window)) return;
                sync_window = value;
                OnPropertyChanged();
            }
        }

        private int scheduler_interval;
        public int SchedulerInterval
        {
            get { return scheduler_interval; }
            set
            {
                if (value == scheduler_interval) return;
                scheduler_interval = value;
                OnPropertyChanged();
            }
        }

        private string calendar_id;
        public string CalendarId
        {
            get { return calendar_id; }
            set
            {
                if (value == calendar_id) return;
                calendar_id = value;
                OnPropertyChanged();
            }
        }

        public AddinSettings()
        {
            IsLoggedIn = false;
            SyncWindow = 30;
            SchedulerInterval = 10;
            CalendarId = string.Empty;
        }

        public void Initialize(string dir)
        {
            base_dir = dir;
            PropertyChanged += (o, a) => Save();
        }

        public static AddinSettings Load(string dir)
        {
            log.Trace("Loading");

            var path = Path.Combine(dir, Filename);
            if (!File.Exists(path))
                return new AddinSettings();

            var json = File.ReadAllText(path);
            var settings = JsonConvert.DeserializeObject<AddinSettings>(json);
            settings.Initialize(dir);
            return settings;
        }

        public void Save()
        {
            log.Trace("Saving");

            var path = Path.Combine(base_dir, Filename);
            var json = JsonConvert.SerializeObject(this, Formatting.Indented);
            File.WriteAllText(path, json);
        }
    }
}
