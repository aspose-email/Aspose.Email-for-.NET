// Demonstrates listing the appointments of a calendar page by page.
//
// ListAppointmentsByPage returns one page of appointments together with TotalCount and
// LastPage; increase the offset until LastPage is true. The example creates eight
// appointments in a uniquely named calendar folder, reads them three at a time, and
// deletes the folder at the end.

using System;
using System.Linq;
using Aspose.Email.Calendar;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class PagingSupportForListingAppointments
    {
        public static void Run()
        {
            const int itemsPerPage = 3;
            var self = ClientBuilder.EwsUserName;
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var calendar = client.CreateFolder(client.MailboxInfo.CalendarUri, folderName,
                    ExchangeFolderType.Appointment);

                try
                {
                    var start = DateTime.Today.AddDays(1).AddHours(9);
                    for (var i = 0; i < 8; i++)
                    {
                        var slotStart = start.AddHours(i);
                        client.CreateAppointment(new Appointment("Room 1", slotStart, slotStart.AddMinutes(30),
                            self, self) { Summary = $"Slot {i + 1}" }, calendar.Uri);
                    }

                    var page = client.ListAppointmentsByPage(calendar.Uri, itemsPerPage);
                    Console.WriteLine($"{page.TotalCount} appointment(s), {itemsPerPage} per page.");

                    while (true)
                    {
                        Console.WriteLine($"Page {page.PageOffset}: " +
                                          string.Join(", ", page.Items.Select(appointment => appointment.Summary)));
                        if (page.LastPage)
                            break;

                        page = client.ListAppointmentsByPage(calendar.Uri, itemsPerPage, page.PageOffset + 1);
                    }
                }
                finally
                {
                    client.DeleteFolder(calendar.Uri, true);
                }
            }
        }
    }
}
