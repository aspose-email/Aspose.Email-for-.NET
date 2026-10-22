// Demonstrates working with a secondary calendar - one besides the default Calendar.
//
// A calendar folder is created with the Appointment folder type. Appointment methods
// accept the folder URI explicitly, or use CurrentCalendarFolderUri as the default when
// none is given - handy when all work happens in one secondary calendar. The example
// creates a uniquely named calendar, adds an event both ways, and deletes the calendar
// at the end.

using System;
using Aspose.Email.Calendar;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class SecondaryCalendarEvents
    {
        public static void Run()
        {
            var self = ClientBuilder.EwsUserName;
            var folderName = "Aspose-Calendar-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var calendar = client.CreateFolder(client.MailboxInfo.CalendarUri, folderName,
                    ExchangeFolderType.Appointment);

                try
                {
                    var start = DateTime.Today.AddDays(2).AddHours(9);

                    // With the folder URI given explicitly.
                    client.CreateAppointment(
                        new Appointment("Room 121", start, start.AddHours(1), self, self)
                        {
                            Summary = "Explicit folder"
                        },
                        calendar.Uri);

                    // With the folder set once as the current calendar.
                    client.CurrentCalendarFolderUri = calendar.Uri;
                    client.CreateAppointment(
                        new Appointment("Room 122", start.AddHours(2), start.AddHours(3), self, self)
                        {
                            Summary = "Current folder"
                        });

                    Console.WriteLine($"Events in '{folderName}':");
                    foreach (var appointment in client.ListAppointments(calendar.Uri))
                        Console.WriteLine($"  {appointment.StartDate:g}  {appointment.Summary}");
                }
                finally
                {
                    client.CurrentCalendarFolderUri = null;
                    client.DeleteFolder(calendar.Uri, true);
                }
            }
        }
    }
}
