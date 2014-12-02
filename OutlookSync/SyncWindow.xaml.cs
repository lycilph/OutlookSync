using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace OutlookSync
{
    public partial class SyncWindow
    {
        private readonly AddinSettings settings;
        private readonly SyncEngine sync_engine;

        public ObservableCollection<string> Messages
        {
            get { return (ObservableCollection<string>)GetValue(MessagesProperty); }
            set { SetValue(MessagesProperty, value); }
        }
        public static readonly DependencyProperty MessagesProperty =
            DependencyProperty.Register("Messages", typeof(ObservableCollection<string>), typeof(SyncWindow), new PropertyMetadata(null));

        public ObservableCollection<StoredAppointment> OutlookAppointments
        {
            get { return (ObservableCollection<StoredAppointment>)GetValue(OutlookAppointmentsProperty); }
            set { SetValue(OutlookAppointmentsProperty, value); }
        }
        public static readonly DependencyProperty OutlookAppointmentsProperty =
            DependencyProperty.Register("OutlookAppointments", typeof(ObservableCollection<StoredAppointment>), typeof(SyncWindow), new PropertyMetadata(null));

        public ObservableCollection<StoredAppointment> GoogleAppointments
        {
            get { return (ObservableCollection<StoredAppointment>)GetValue(GoogleAppointmentsProperty); }
            set { SetValue(GoogleAppointmentsProperty, value); }
        }
        public static readonly DependencyProperty GoogleAppointmentsProperty =
            DependencyProperty.Register("GoogleAppointments", typeof(ObservableCollection<StoredAppointment>), typeof(SyncWindow), new PropertyMetadata(null));

        public ObservableCollection<StoredAppointment> AppointmentsToRemove
        {
            get { return (ObservableCollection<StoredAppointment>)GetValue(AppointmentsToRemoveProperty); }
            set { SetValue(AppointmentsToRemoveProperty, value); }
        }
        public static readonly DependencyProperty AppointmentsToRemoveProperty =
            DependencyProperty.Register("AppointmentsToRemove", typeof(ObservableCollection<StoredAppointment>), typeof(SyncWindow), new PropertyMetadata(null));

        public ObservableCollection<StoredAppointment> AppointmentsToAdd
        {
            get { return (ObservableCollection<StoredAppointment>)GetValue(AppointmentsToAddProperty); }
            set { SetValue(AppointmentsToAddProperty, value); }
        }
        public static readonly DependencyProperty AppointmentsToAddProperty =
            DependencyProperty.Register("AppointmentsToAdd", typeof(ObservableCollection<StoredAppointment>), typeof(SyncWindow), new PropertyMetadata(null));
        
        public SyncWindow()
        {
            InitializeComponent();
            DataContext = this;

            settings = Globals.ThisAddIn.Settings;
            sync_engine = Globals.ThisAddIn.SyncEngine;

            Messages = new ObservableCollection<string>();
            OutlookAppointments = new ObservableCollection<StoredAppointment>();
            GoogleAppointments = new ObservableCollection<StoredAppointment>();
            AppointmentsToRemove = new ObservableCollection<StoredAppointment>();
            AppointmentsToAdd = new ObservableCollection<StoredAppointment>();

            Loaded += OnLoaded;
        }

        private async void OnLoaded(object sender, RoutedEventArgs routed_event_args)
        {
            Loaded -= OnLoaded;

            await Analyze();
        }

        private void FindItemsToRemove()
        {
            Messages.Add("Finding appointments to remove");

            var items = GoogleAppointments.Except(OutlookAppointments).ToList();
            AppointmentsToRemove = new ObservableCollection<StoredAppointment>(items);

            Messages.Add("Found " + items.Count + " item(s)");
        }

        private void FindItemsToAdd()
        {
            Messages.Add("Finding appointments to add");

            var items = OutlookAppointments.Except(GoogleAppointments).ToList();
            AppointmentsToAdd = new ObservableCollection<StoredAppointment>(items);

            Messages.Add("Found " + items.Count + " item(s)");
        }

        private async Task LoadOutlookAppointments()
        {
            Messages.Add("Loading outlook appointments");

            var start = DateTime.Now.Date;
            var end = start.AddDays(settings.SyncWindow);
            var items = await Task.Factory.StartNew(() => sync_engine.GetOutlookItems(start, end));
            OutlookAppointments = new ObservableCollection<StoredAppointment>(items);

            Messages.Add("Found " + items.Count + " appointments");
        }

        private async Task LoadGoogleAppointments()
        {
            Messages.Add("Loading google appointments");

            var start = DateTime.Now.Date;
            var end = start.AddDays(settings.SyncWindow);
            var items = await Task.Factory.StartNew(() => sync_engine.GetGoogleItems(settings.CalendarId, start, end));
            GoogleAppointments = new ObservableCollection<StoredAppointment>(items);

            Messages.Add("Found " + items.Count + " appointments");
        }

        private async void ExecuteClick(object sender, RoutedEventArgs e)
        {
            await Execute();
        }

        private async Task Execute()
        {
            ExecuteButton.IsEnabled = false;
            AnalyzeButton.IsEnabled = false;

            if (AppointmentsToRemove.Any())
            {
                var items_to_remove = AppointmentsToRemove.ToList();
                AppointmentsToRemove.Clear();

                Messages.Add("Removing items");
                await Task.Factory.StartNew(() => sync_engine.RemoveGoogleItems(settings.CalendarId, items_to_remove));                
            }

            if (AppointmentsToAdd.Any())
            {
                var items_to_add = AppointmentsToAdd.ToList();
                AppointmentsToAdd.Clear();

                Messages.Add("Adding items");
                await Task.Factory.StartNew(() => sync_engine.AddGoogleItems(settings.CalendarId, items_to_add));                
            }

            Messages.Add("Execute done");

            ExecuteButton.IsEnabled = true;
            AnalyzeButton.IsEnabled = true;
        }

        private async void AnalyzeClick(object sender, RoutedEventArgs e)
        {
            await Analyze();
        }

        private async Task Analyze()
        {
            ExecuteButton.IsEnabled = false;
            AnalyzeButton.IsEnabled = false;

            Messages.Clear();
            OutlookAppointments.Clear();
            GoogleAppointments.Clear();
            AppointmentsToAdd.Clear();
            AppointmentsToRemove.Clear();

            await LoadOutlookAppointments();
            await LoadGoogleAppointments();

            FindItemsToAdd();
            FindItemsToRemove();

            Messages.Add("Analyze done");

            ExecuteButton.IsEnabled = true;
            AnalyzeButton.IsEnabled = true;
        }
    }
}
