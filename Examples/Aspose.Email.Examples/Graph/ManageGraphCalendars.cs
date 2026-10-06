// Demonstrates the calendar side of Graph: listing the calendars of a mailbox and
// creating, reading and updating the appointments in one.

using System;
using System.Linq;
using Aspose.Email.Clients.Graph;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Graph
{
    internal static class ManageGraphCalendars
    {
        public static void Run()
        {
            if (!ClientBuilder.IsGraphConfigured)
            {
                GraphExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Graph(AuthType.ModernWithAppPermission))
            {
                var calendars = client.ListCalendars(null);
                Console.WriteLine($"{calendars.Count} calendar(s):");

                foreach (var calendar in calendars)
                    Console.WriteLine($"  {calendar.Name} ({calendar.Color}), default: {calendar.IsDefaultCalendar}, id: {calendar.ItemId}");

                var primary = calendars.FirstOrDefault();
                if (primary == null)
                {
                    Console.WriteLine("The mailbox has no calendar.");
                    return;
                }

                var start = DateTime.UtcNow.Date.AddDays(1).AddHours(10);

                var appointment = new MapiCalendar(
                    "Conference Room",
                    "Team Meeting",
                    "Discuss project status and updates.",
                    start,
                    start.AddHours(1));

                var created = client.CreateCalendarItem(primary.ItemId, appointment);
                Console.WriteLine($"\nCreated: {created.Subject} at {created.StartDate:u}");

                var items = client.ListCalendarItems(primary.ItemId, null);
                Console.WriteLine($"{items.Count} item(s) on {primary.Name}");

                var fetched = client.FetchCalendarItem(created.ItemId);
                Console.WriteLine($"Fetched: {fetched.Subject}, location: {fetched.Location}");

                fetched.Location = "Zoom Meeting";

                // SkipAttachments keeps the update from re-uploading anything attached.
                var updated = client.UpdateCalendarItem(fetched, new UpdateSettings { SkipAttachments = true });
                Console.WriteLine($"Updated location to: {updated.Location}");

                client.Delete(created.ItemId);
                Console.WriteLine("Deleted the appointment.");
            }
        }
    }
}
