using System;
using System.IO;
using Newtonsoft.Json;

namespace OutlookSync
{
    public class AddinSettings : ObservableObject
    {
        private const string Filename = "settings.json";

        private string dir;

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

        public AddinSettings() : this(string.Empty) {}
        public AddinSettings(string dir)
        {
            this.dir = dir;
            IsLoggedIn = false;
            SyncWindow = 30;
            CalendarId = string.Empty;

            PropertyChanged += (o, a) => Save();
        }

        public static AddinSettings Load(string dir)
        {
            var path = Path.Combine(dir, Filename);

            if (!File.Exists(path))
                return new AddinSettings(dir);

            var json = File.ReadAllText(path);
            var settings = JsonConvert.DeserializeObject<AddinSettings>(json);
            settings.dir = dir;
            return settings;
        }

        public void Save()
        {
            var path = Path.Combine(dir, Filename);
            var json = JsonConvert.SerializeObject(this, Formatting.Indented);
            File.WriteAllText(path, json);
        }
    }
}
