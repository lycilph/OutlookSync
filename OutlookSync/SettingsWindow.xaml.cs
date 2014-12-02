using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;

namespace OutlookSync
{
    public partial class SettingsWindow
    {
        private readonly AddinSettings settings;
        private readonly SyncEngine sync_engine;

        public int SyncWindow
        {
            get { return (int)GetValue(SyncWindowProperty); }
            set { SetValue(SyncWindowProperty, value); }
        }
        public static readonly DependencyProperty SyncWindowProperty =
            DependencyProperty.Register("SyncWindow", typeof(int), typeof(SettingsWindow), new PropertyMetadata(0));

        public ObservableCollection<GoogleCalendar> Calendars
        {
            get { return (ObservableCollection<GoogleCalendar>)GetValue(CalendarsProperty); }
            set { SetValue(CalendarsProperty, value); }
        }
        public static readonly DependencyProperty CalendarsProperty =
            DependencyProperty.Register("Calendars", typeof(ObservableCollection<GoogleCalendar>), typeof(SettingsWindow), new PropertyMetadata(null));

        public SettingsWindow()
        {
            InitializeComponent();
            DataContext = this;

            Loaded += OnLoaded;

            settings = Globals.ThisAddIn.Settings;
            sync_engine = Globals.ThisAddIn.SyncEngine;

            SyncWindow = settings.SyncWindow;
        }

        private async void OnLoaded(object sender, RoutedEventArgs routed_event_args)
        {
            Loaded -= OnLoaded;

            if (settings.IsLoggedIn)
            {
                await LoadCalendars();

                SyncWindowTextBox.IsEnabled = true;
                CalendarsComboBox.IsEnabled = true;
                OkButton.IsEnabled = true;
            }
            else
            {
                LoginButton.IsEnabled = true;
            }
        }

        private void OnOkClick(object sender, RoutedEventArgs e)
        {
            if (SyncWindowProperty.IsValidValue(SyncWindow))
                settings.SyncWindow = SyncWindow;

            var view = CollectionViewSource.GetDefaultView(Calendars);
            var current = view.CurrentItem as GoogleCalendar;
            if (current != null)
                settings.CalendarId = current.Id;

            Close();
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private async void OnLoginClick(object sender, RoutedEventArgs e)
        {
            await Task.Factory.StartNew(sync_engine.Initialize);
            await LoadCalendars();

            settings.IsLoggedIn = true;

            LoginButton.IsEnabled = false;
            SyncWindowTextBox.IsEnabled = true;
            CalendarsComboBox.IsEnabled = true;
            OkButton.IsEnabled = true;
        }

        private async Task LoadCalendars()
        {
            var list = await Task.Factory.StartNew(() => sync_engine.GetGoogleCalendars());
            Calendars = new ObservableCollection<GoogleCalendar>(list);

            if (string.IsNullOrWhiteSpace(settings.CalendarId))
                return;

            var calendar = Calendars.SingleOrDefault(c => c.Id == settings.CalendarId);
            if (calendar == null)
                return;

            var view = CollectionViewSource.GetDefaultView(Calendars);
            view.MoveCurrentTo(calendar);
        }
    }
}
