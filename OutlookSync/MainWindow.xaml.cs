using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using Newtonsoft.Json;

namespace OutlookSync
{
    public partial class MainWindow
    {
        private readonly SyncEngine sync_engine;

        public ObservableCollection<StoredAppointment> CacheItems
        {
            get { return (ObservableCollection<StoredAppointment>)GetValue(CacheItemsProperty); }
            set { SetValue(CacheItemsProperty, value); }
        }
        public static readonly DependencyProperty CacheItemsProperty =
            DependencyProperty.Register("CacheItems", typeof(ObservableCollection<StoredAppointment>), typeof(MainWindow), new PropertyMetadata(null));

        public DateTime OutlookStart
        {
            get { return (DateTime)GetValue(OutlookStartProperty); }
            set { SetValue(OutlookStartProperty, value); }
        }
        public static readonly DependencyProperty OutlookStartProperty =
            DependencyProperty.Register("OutlookStart", typeof(DateTime), typeof(MainWindow), new PropertyMetadata(DateTime.Now));

        public DateTime OutlookEnd
        {
            get { return (DateTime)GetValue(OutlookEndProperty); }
            set { SetValue(OutlookEndProperty, value); }
        }
        public static readonly DependencyProperty OutlookEndProperty =
            DependencyProperty.Register("OutlookEnd", typeof(DateTime), typeof(MainWindow), new PropertyMetadata(DateTime.Now.AddMonths(1)));

        public string OutlookStatus
        {
            get { return (string)GetValue(OutlookStatusProperty); }
            set { SetValue(OutlookStatusProperty, value); }
        }
        public static readonly DependencyProperty OutlookStatusProperty =
            DependencyProperty.Register("OutlookStatus", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));

        public ObservableCollection<StoredAppointment> OutlookItems
        {
            get { return (ObservableCollection<StoredAppointment>)GetValue(OutlookItemsProperty); }
            set { SetValue(OutlookItemsProperty, value); }
        }
        public static readonly DependencyProperty OutlookItemsProperty =
            DependencyProperty.Register("OutlookItems", typeof(ObservableCollection<StoredAppointment>), typeof(MainWindow), new PropertyMetadata(null));

        public DateTime GoogleStart
        {
            get { return (DateTime)GetValue(GoogleStartProperty); }
            set { SetValue(GoogleStartProperty, value); }
        }
        public static readonly DependencyProperty GoogleStartProperty =
            DependencyProperty.Register("GoogleStart", typeof(DateTime), typeof(MainWindow), new PropertyMetadata(DateTime.Now));

        public DateTime GoogleEnd
        {
            get { return (DateTime)GetValue(GoogleEndProperty); }
            set { SetValue(GoogleEndProperty, value); }
        }
        public static readonly DependencyProperty GoogleEndProperty =
            DependencyProperty.Register("GoogleEnd", typeof(DateTime), typeof(MainWindow), new PropertyMetadata(DateTime.Now.AddMonths(1)));

        public string GoogleStatus
        {
            get { return (string)GetValue(GoogleStatusProperty); }
            set { SetValue(GoogleStatusProperty, value); }
        }
        public static readonly DependencyProperty GoogleStatusProperty =
            DependencyProperty.Register("GoogleStatus", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));

        public ObservableCollection<StoredAppointment> GoogleItems
        {
            get { return (ObservableCollection<StoredAppointment>)GetValue(GoogleItemsProperty); }
            set { SetValue(GoogleItemsProperty, value); }
        }
        public static readonly DependencyProperty GoogleItemsProperty =
            DependencyProperty.Register("GoogleItems", typeof(ObservableCollection<StoredAppointment>), typeof(MainWindow), new PropertyMetadata(null));

        public ObservableCollection<string> Calendars
        {
            get { return (ObservableCollection<string>)GetValue(CalendarsProperty); }
            set { SetValue(CalendarsProperty, value); }
        }
        public static readonly DependencyProperty CalendarsProperty =
            DependencyProperty.Register("Calendars", typeof(ObservableCollection<string>), typeof(MainWindow), new PropertyMetadata(null));

        public MainWindow()
        {
            InitializeComponent();

            DataContext = this;
            Loaded += OnLoaded;

            sync_engine = new SyncEngine();

            OutlookItems = new ObservableCollection<StoredAppointment>();
            GoogleItems = new ObservableCollection<StoredAppointment>();
            CacheItems = new ObservableCollection<StoredAppointment>();
            Calendars = new ObservableCollection<string>();
        }

        private void OnLoaded(object sender, RoutedEventArgs routedEventArgs)
        {
            Mouse.OverrideCursor = Cursors.Wait;
            sync_engine.Initialize();
            
            Calendars = new ObservableCollection<string>(sync_engine.GetGoogleCalendars());
            var view = CollectionViewSource.GetDefaultView(Calendars);
            view.MoveCurrentToFirst();

            Mouse.OverrideCursor = null;
        }

        private void OutlookGetClick(object sender, RoutedEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Wait;
            
            var items = sync_engine.GetOutlookItems(OutlookStart, OutlookEnd);
            OutlookItems = new ObservableCollection<StoredAppointment>(items);
            OutlookStatus = string.Format("{0} item(s) found", OutlookItems.Count);

            MainTabControl.SelectedIndex = 1;
            Mouse.OverrideCursor = null;
        }

        private void GoogleGetClick(object sender, RoutedEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Wait;

            var view = CollectionViewSource.GetDefaultView(Calendars);
            var item = view.CurrentItem as string;
            if (item == null)
            {
                Mouse.OverrideCursor = null;
                return;
            }
            
            var elements = item.Split(new[] {':'}, StringSplitOptions.None);
            var calendar_id = elements[1];

            var items = sync_engine.GetGoogleItems(calendar_id, GoogleStart, GoogleEnd);
            GoogleItems = new ObservableCollection<StoredAppointment>(items);
            GoogleStatus = string.Format("{0} item(s) found", GoogleItems.Count);

            MainTabControl.SelectedIndex = 1;
            Mouse.OverrideCursor = null;
        }

        private void LoadClick(object sender, RoutedEventArgs e)
        {
            var file = Path.Combine(sync_engine.BaseDir, "cache_items.json");
            if (File.Exists(file))
                return;

            var json = File.ReadAllText(file);
            var items = JsonConvert.DeserializeObject<List<StoredAppointment>>(json);
            CacheItems = new ObservableCollection<StoredAppointment>(items);
        }

        private void SaveClick(object sender, RoutedEventArgs e)
        {
            var file = Path.Combine(sync_engine.BaseDir, "cache_items.json");
            var json = JsonConvert.SerializeObject(CacheItems, Formatting.Indented);
            File.WriteAllText(file,json);
        }

        private void CacheOutlookItemsClick(object sender, RoutedEventArgs e)
        {
            CacheItems = new ObservableCollection<StoredAppointment>(OutlookItems);
        }

        private void CacheGoogleItemsClick(object sender, RoutedEventArgs e)
        {
            CacheItems = new ObservableCollection<StoredAppointment>(GoogleItems);
        }

        private void GetOutlookItems(object sender, RoutedEventArgs e)
        {
            

        }
    }
}
