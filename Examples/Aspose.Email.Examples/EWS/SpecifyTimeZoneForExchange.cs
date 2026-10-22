// Demonstrates choosing the time zone in which task and calendar dates are returned.
//
// Exchange stores dates in UTC and converts them for the client. TimezoneId takes a
// Windows time zone id (see GetServerTimeZonesUsingEWS for the ids the server knows);
// dates read after setting it are expressed in that zone. The example lists the same
// tasks in two time zones.

using System;
using System.Linq;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class SpecifyTimeZoneForExchange
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                foreach (var timeZone in new[] { "UTC", "Tokyo Standard Time" })
                {
                    client.TimezoneId = timeZone;

                    Console.WriteLine($"Tasks in {client.TimezoneId}:");
                    foreach (var task in client.ListTasks(client.MailboxInfo.TasksUri, 3).Cast<ExchangeTask>())
                        Console.WriteLine($"  {task.Subject}: start {task.StartDate:g}, due {task.DueDate:g}");
                    Console.WriteLine();
                }
            }
        }
    }
}
