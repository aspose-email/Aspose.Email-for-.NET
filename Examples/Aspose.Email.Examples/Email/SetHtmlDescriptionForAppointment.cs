// Demonstrates how to give an appointment a formatted description. HtmlDescription
// is written to the X-ALT-DESC header, and Description keeps the plain-text version
// for clients that cannot render HTML.

using System;
using System.IO;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.Email
{
    internal static class SetHtmlDescriptionForAppointment
    {
        public static void Run()
        {
            var appointment = new Appointment("Bygget 83",
                DateTime.UtcNow,
                DateTime.UtcNow.AddHours(1),
                new MailAddress("TintinStrom@from.com", "Tintin Strom"),
                new MailAddress("AinaMartensson@to.com", "Aina Martensson"))
            {
                Summary = "Party",
                Description = "Hi, I'm happy to invite you to our party.",
                HtmlDescription = @"
    <html>
     <style type=""text/css"">
      .text {
             font-family:'Comic Sans MS';
             font-size:16px;
            }
     </style>
    <body>
     <p class=""text"">Hi, I'm happy to invite you to our party.</p>
    </body>
    </html>"
            };

            var outputPath = Data.Out/"SetHtmlDescriptionForAppointment_out.ics";
            appointment.Save(outputPath, AppointmentSaveFormat.Ics);

            Console.WriteLine($"Appointment: {appointment.Summary}");

            foreach (var line in File.ReadLines(outputPath))
            {
                if (line.StartsWith("X-ALT-DESC", StringComparison.Ordinal) ||
                    line.StartsWith("DESCRIPTION", StringComparison.Ordinal))
                {
                    Console.WriteLine($"  {(line.Length <= 120 ? line : line.Substring(0, 120) + "...")}");
                }
            }

            Console.WriteLine($"\nSaved to {outputPath}");
        }
    }
}
