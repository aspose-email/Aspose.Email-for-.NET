// Demonstrates the time zones the Exchange server knows.
//
// Exchange identifies time zones by Windows ids such as "W. Europe Standard Time".
// GetServerTimeZoneIds lists them all, or - given candidate ids - returns the ones the
// server recognises, which is a quick way to validate an id before assigning it to
// TimezoneId for calendar and task operations.

using System;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    internal static class GetServerTimeZonesUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var all = client.GetServerTimeZoneIds();
                Console.WriteLine($"The server knows {all.Length} time zone(s), for example:");
                foreach (var id in all.Take(10))
                    Console.WriteLine("  " + id);

                var known = client.GetServerTimeZoneIds("W. Europe Standard Time", "Mars Standard Time");
                Console.WriteLine($"\nOf 'W. Europe Standard Time' and 'Mars Standard Time' it recognises: " +
                                  string.Join(", ", known));

                Console.WriteLine($"\nTime zone the client uses: {client.TimezoneId ?? "(server default)"}");
            }
        }
    }
}
