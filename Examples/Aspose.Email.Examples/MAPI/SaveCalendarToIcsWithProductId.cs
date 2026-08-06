// Demonstrates the two ICS save options that control what a calendar item looks like
// on the way out: the PRODID line, and whether the original DTSTAMP is preserved.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SaveCalendarToIcsWithProductId
    {
        public static void Run()
        {
            var msg = MapiMessage.Load(Data.Mapi/"Test Meeting.msg");
            var calendar = (MapiCalendar)msg.ToMapiMessageItem();

            var options = new MapiCalendarIcsSaveOptions
            {
                // Without this the timestamp is refreshed to the moment of saving.
                KeepOriginalDateTimeStamp = true,
                ProductIdentifier = "Foo Ltd"
            };

            var outputPath = Data.Out/"SaveCalendarToIcsWithProductId_out.ics";
            calendar.Save(outputPath, options);

            Console.WriteLine($"Appointment: {calendar.Subject}");
            Console.WriteLine($"PRODID:      {options.ProductIdentifier}");

            foreach (var line in System.IO.File.ReadLines(outputPath))
            {
                if (line.StartsWith("PRODID", StringComparison.Ordinal) ||
                    line.StartsWith("DTSTAMP", StringComparison.Ordinal))
                {
                    Console.WriteLine($"  {line}");
                }
            }

            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
