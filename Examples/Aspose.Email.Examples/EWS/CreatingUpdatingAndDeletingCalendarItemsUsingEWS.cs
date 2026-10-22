// Demonstrates the life cycle of a calendar item: create, read, update, cancel.
//
// CreateAppointment stores an Appointment in a calendar folder and returns its id;
// FetchAppointment reads it back, UpdateAppointment saves changes, and
// CancelAppointment cancels it. The configured mailbox is both organiser and attendee,
// so any meeting messages stay in that mailbox. The example works in a uniquely named
// calendar folder and deletes it at the end.

using System;
using Aspose.Email.Calendar;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class CreatingUpdatingAndDeletingCalendarItemsUsingEWS
    {
        public static void Run()
        {
            var self = ClientBuilder.EwsUserName;
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var calendar = client.CreateFolder(client.MailboxInfo.CalendarUri, folderName,
                    ExchangeFolderType.Appointment);

                try
                {
                    var start = DateTime.Today.AddDays(1).AddHours(10);
                    var appointment = new Appointment("Room 112", start, start.AddHours(1), self, self)
                    {
                        Summary = "Planning session",
                        Description = "Plan the next sprint."
                    };

                    var id = client.CreateAppointment(appointment, calendar.Uri);
                    var created = client.FetchAppointment(id, calendar.Uri);
                    Console.WriteLine($"Created:   {created.Summary} in {created.Location} at {created.StartDate:g}");

                    created.Location = "Room 115";
                    created.Summary = "Planning session (moved)";
                    client.UpdateAppointment(created, calendar.Uri);

                    var updated = client.FetchAppointment(id, calendar.Uri);
                    Console.WriteLine($"Updated:   {updated.Summary} in {updated.Location}");

                    client.CancelAppointment(id, calendar.Uri);
                    Console.WriteLine($"Cancelled; appointments left: {client.ListAppointments(calendar.Uri).Length}");
                }
                finally
                {
                    client.DeleteFolder(calendar.Uri, true);
                }
            }
        }
    }
}
