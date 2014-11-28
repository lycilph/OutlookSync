using System;
using Google.Apis.Calendar.v3.Data;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSync
{
    public class StoredAppointment : IEquatable<StoredAppointment>
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public string Location { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public DateTime LastModificationTime { get; set; }

        public StoredAppointment() { }

        public StoredAppointment(Outlook.AppointmentItem item)
        {
            Subject = item.Subject;
            Location = item.Location;
            Start = item.Start;
            End = item.End;
            LastModificationTime = item.LastModificationTime;
            Id = item.GlobalAppointmentID + "-" + Start.ToShortDateString() + "-" + End.ToShortDateString();
        }

        public StoredAppointment(Event e)
        {
            Id = e.Id;
            Subject = e.Summary;
            Location = e.Location;

            if (e.Start.Date != null)
                Start = DateTime.Parse(e.Start.Date);

            if (e.End.Date != null)
                End = DateTime.Parse(e.End.Date);

            if (e.Updated != null)
                LastModificationTime = e.Updated.Value;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            return obj.GetType() == GetType() &&
                   Equals((StoredAppointment) obj);
        }

        public bool Equals(StoredAppointment other)
        {
            if (ReferenceEquals(null, other)) return false;
            if (ReferenceEquals(this, other)) return true;
            return string.Equals(Subject, other.Subject) && 
                  string.Equals(Location, other.Location) && 
                  Start.Equals(other.Start) && 
                  End.Equals(other.End);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hash_code = (Subject != null ? Subject.GetHashCode() : 0);
                hash_code = (hash_code * 397) ^ (Location != null ? Location.GetHashCode() : 0);
                hash_code = (hash_code * 397) ^ Start.GetHashCode();
                hash_code = (hash_code * 397) ^ End.GetHashCode();
                return hash_code;
            }
        }

        public static bool operator ==(StoredAppointment left, StoredAppointment right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(StoredAppointment left, StoredAppointment right)
        {
            return !Equals(left, right);
        }
    }
}
