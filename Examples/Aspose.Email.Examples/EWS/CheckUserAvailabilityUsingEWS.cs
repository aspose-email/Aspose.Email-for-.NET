// Demonstrates checking free/busy information before scheduling a meeting.
//
// CheckUserAvailability returns, for each user and time window, the busy periods in
// their calendar (with subject and location if they share them) and their working
// hours - what the Outlook scheduling assistant shows. Put real addresses of your
// organisation into the list below.

using System;
using System.Collections.Specialized;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class CheckUserAvailabilityUsingEWS
    {
        public static void Run()
        {
            var users = new StringCollection { "user1@example.com", "user2@example.com" };

            var tomorrow = DateTime.Today.AddDays(1);
            var window = new DateRange(tomorrow.AddHours(8), tomorrow.AddHours(18));

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var availability = client.CheckUserAvailability(users, window);

                foreach (var user in availability)
                {
                    var hours = user.GetWorkingHours(tomorrow);
                    Console.WriteLine($"{user.User}, working hours tomorrow: {hours?.StartTime:t}-{hours?.EndTime:t}");

                    if (user.CalendarAppointments.Count == 0)
                        Console.WriteLine("  free all day");

                    foreach (var busy in user.CalendarAppointments)
                    {
                        Console.WriteLine($"  {busy.Duration.StartTime:t}-{busy.Duration.EndTime:t}  {busy.Busyness}" +
                                          (string.IsNullOrEmpty(busy.Subject) ? "" : $"  {busy.Subject}"));
                    }
                }
            }
        }
    }
}
