// Demonstrates reading custom (named) MAPI properties of calendar items.
//
// FetchMapiCalendar returns calendar items as MapiCalendar objects; extended properties
// are included only when their descriptors are passed in. The example creates an
// appointment carrying a custom property in a uniquely named calendar folder, reads the
// property back and deletes the folder at the end.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.EWS
{
    internal static class RetreiveExtAttributesForCalendarItems
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            var property = new PidNamePropertyDescriptor("ProjectCode", PropertyDataType.String,
                KnownPropertySets.PublicStrings);

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var calendar = client.CreateFolder(client.MailboxInfo.CalendarUri, folderName,
                    ExchangeFolderType.Appointment);

                try
                {
                    var start = DateTime.Today.AddDays(1).AddHours(14);
                    var item = new MapiCalendar("Room 7", "Project review", "Quarterly review",
                        start, start.AddHours(1));
                    item.SetProperty(property, "PRJ-42");

                    client.CreateAppointment(item, calendar.Uri, true);

                    var uris = client.ListItems(calendar.Uri);
                    foreach (var fetched in client.FetchMapiCalendar(uris, new PropertyDescriptor[] { property }))
                    {
                        var value = fetched.NamedProperties.Values.FirstOrDefault()?.GetString();
                        Console.WriteLine($"{fetched.Subject}: ProjectCode = {value}");
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
