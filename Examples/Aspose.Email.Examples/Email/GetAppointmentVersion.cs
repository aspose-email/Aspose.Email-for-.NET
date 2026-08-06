// Demonstrates how to read the iCalendar version, method and product id of a
// calendar file, which is what decides how the rest of it should be interpreted.

using System;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.Email
{
    internal static class GetAppointmentVersion
    {
        public static void Run()
        {
            foreach (var fileName in new[] { "test.ics", "US-Holidays.ics" })
            {
                var reader = new CalendarReader(Data.Email/fileName);
                Console.WriteLine($"{fileName}: version {reader.Version}, method {reader.Method}, " +
                                  $"{reader.Count} event(s)");

                var app = Appointment.Load(Data.Email/fileName);
                Console.WriteLine($"  Appointment.Version:   {app.Version}");
                Console.WriteLine($"  Appointment.ProductId: {app.ProductId}");

                if (app.Version == "2.0")
                    Console.WriteLine("  (RFC 5545 iCalendar)");
            }
        }
    }
}
