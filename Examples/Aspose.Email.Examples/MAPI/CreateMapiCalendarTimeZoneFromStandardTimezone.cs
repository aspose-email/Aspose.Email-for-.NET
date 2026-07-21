// Demonstrates how to assign a MapiCalendarTimeZone (from the system local time zone) to a calendar item.

using System;
using Aspose.Email.Mapi;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateMapiCalendarTimeZoneFromStandardTimezone
    {
        public static void Run()
        {
            var message = MapiMessage.Load(Data.Mapi/"Test Meeting.msg");
            var calendar = (MapiCalendar)message.ToMapiMessageItem();

            // MapiCalendarTimeZone can be built straight from a .NET TimeZoneInfo, so the
            // transition rules do not have to be described by hand.
            var timeZone = new MapiCalendarTimeZone(TimeZoneInfo.Local);
            calendar.StartDateTimeZone = timeZone;
            calendar.EndDateTimeZone = timeZone;

            var outputPath = Data.Out/"MapiCalendarWithTimezone_out.ics";
            calendar.Save(outputPath, AppointmentSaveFormat.Ics);

            Console.WriteLine($"Appointment: {calendar.Subject}");
            Console.WriteLine($"Time zone:   {timeZone.KeyName}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
