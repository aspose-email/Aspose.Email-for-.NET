// Demonstrates how to set the state of a calendar item explicitly. Normally the
// state follows from how the item was built; SetStateForced overrides that, which
// is what turns a plain appointment into a received meeting.

using System;
using Aspose.Email.Calendar;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class ForceCalendarState
    {
        public static void Run()
        {
            var appointment = new MapiCalendar(
                "LAKE ARGYLE WA 6743",
                "Appointment",
                "This is a very important meeting",
                new DateTime(2024, 5, 10, 12, 30, 0, DateTimeKind.Utc),
                new DateTime(2024, 5, 10, 13, 30, 0, DateTimeKind.Utc))
            {
                Organizer = new MapiElectronicAddress
                {
                    EmailAddress = "test@example.com",
                    DisplayName = "Test Display Name"
                }
            };

            // The state is not exposed as a property, so read the underlying MAPI value.
            Console.WriteLine($"State as built: {ReadState(appointment)}");

            appointment.SetStateForced(MapiCalendarState.Meeting | MapiCalendarState.Received);
            Console.WriteLine($"State forced:   {ReadState(appointment)}");

            var outputPath = Data.Out/"ForceCalendarState_out.msg";
            appointment.Save(outputPath, AppointmentSaveFormat.Msg);
            Console.WriteLine($"\nSaved to {outputPath}");
        }

        private static MapiCalendarState ReadState(MapiCalendar appointment)
        {
            var property = appointment.GetUnderlyingMessage()
                .GetProperty(KnownPropertyList.AppointmentStateFlags);

            // Zero is the "plain appointment" state - the enum has no name for it.
            return (MapiCalendarState)(property?.GetInt32() ?? 0);
        }
    }
}
