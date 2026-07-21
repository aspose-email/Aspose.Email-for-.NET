// Demonstrates how to set a custom product identifier (PRODID) when saving
// an appointment as an ICS file.

using Aspose.Email.Calendar;
using System;

namespace Aspose.Email.Examples.Email
{
    internal static class ChangeProdIdOfIcs
    {
        public static void Run()
        {
            var app = new Appointment(
                "location",
                "test appointment",
                "Test Description",
                DateTime.Today,
                DateTime.Today.AddDays(1),
                "first@test.com",
                "second@test.com");

            // PRODID identifies the software that produced the file; it defaults to
            // Aspose.Email unless overridden here.
            var saveOptions = AppointmentIcsSaveOptions.Default;
            saveOptions.ProductId = "Test Corporation";

            var outputPath = Data.Out/"ChangeProdIdOfICS.ics";
            app.Save(outputPath, saveOptions);

            Console.WriteLine($"PRODID set to: {saveOptions.ProductId}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
