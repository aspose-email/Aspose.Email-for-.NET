// Demonstrates how to read multiple calendar events from an ICS file using
// CalendarReader, which streams them one at a time.

using System;
using System.Collections.Generic;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.Email
{
    internal static class ReadMultipleEventsFromIcs
    {
        public static void Run()
        {
            var appointments = new List<Appointment>();

            // Appointment.Load() reads a single event; a file holding several needs the
            // reader, which advances through them one by one.
            var reader = new CalendarReader(Data.Email/"US-Holidays.ics");
            while (reader.NextEvent())
                appointments.Add(reader.Current);

            foreach (var appointment in appointments)
                Console.WriteLine($"{appointment.StartDate:d}  {appointment.Summary}");

            Console.WriteLine($"\nRead {appointments.Count} event(s).");
        }
    }
}
