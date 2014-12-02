namespace OutlookSync
{
    public class GoogleCalendar
    {
        public string DisplayName { get; set; }
        public string Id { get; set; }

        public GoogleCalendar(string name, string id)
        {
            DisplayName = name;
            Id = id;
        }
    }
}
