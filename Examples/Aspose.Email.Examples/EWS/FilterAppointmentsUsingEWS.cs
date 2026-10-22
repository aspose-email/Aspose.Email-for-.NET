// Demonstrates filtering calendar items on the server.
//
// ExchangeQueryBuilder.Appointment offers conditions on start and end time, recurrence
// and more; ListAppointments with the query returns only the matching appointments.
// The example looks at the default calendar.

using System;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class FilterAppointmentsUsingEWS
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var builder = new ExchangeQueryBuilder();
                builder.Appointment.Start.Since(DateTime.Today);
                builder.Appointment.End.BeforeOrEqual(DateTime.Today.AddDays(7));

                var nextWeek = client.ListAppointments(builder.GetQuery());
                Console.WriteLine($"In the next 7 days: {nextWeek.Length} appointment(s)");
                foreach (var appointment in nextWeek)
                    Console.WriteLine($"  {appointment.StartDate:g}  {appointment.Summary}");

                builder = new ExchangeQueryBuilder();
                builder.Appointment.IsRecurring.Equals(true);
                Console.WriteLine($"\nRecurring series: {client.ListAppointments(builder.GetQuery()).Length}");
            }
        }
    }
}
