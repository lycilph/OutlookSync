using System;
using System.IO;
using Newtonsoft.Json;

namespace OutlookSync
{
    public class AddinSettings : ObservableObject
    {
        private const string Filename = "settings.json";

        private string dir;

        private bool is_initialized;
        public bool IsInitialized
        {
            get { return is_initialized; }
            set
            {
                if (value.Equals(is_initialized)) return;
                is_initialized = value;
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
            IsInitialized = false;
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
