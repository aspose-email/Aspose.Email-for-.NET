using System;
using Aspose.Email.Mime;
using Aspose.Email.Calendar;

namespace Aspose.Email.Examples.CSharp.Email.SMTP
{
    class AppointmentInICSFormat
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "test.ics";


            // ExStart:CreateAppointment

            // Create and initialize an instance of the Appointment class
            Appointment appointment = new Appointment(
                "Meeting Room 3 at Office Headquarters",// Location
                "Monthly Meeting",                      // Summary
                "Please confirm your availability.",    // Description
                new DateTime(2015, 2, 8, 13, 0, 0),     // Start date
                new DateTime(2015, 2, 8, 14, 0, 0),     // End date
                "from@domain.com",                      // Organizer
                "attendees@domain.com");                // Attendees

            // Save the appointment to disk in ICS format
            appointment.Save(dstEmail, AppointmentSaveFormat.Ics);
            Console.WriteLine("Appointment created and saved to disk successfully.");
            // ExEnd:CreateAppointment

            // ExStart:LoadAppointment
            // Load an Appointment just created and saved to disk and display its details.
            Appointment loadedAppointment = Appointment.Load(dstEmail);
            Console.WriteLine(Environment.NewLine + "Loaded Appointment details are as follows:");
            // Display the appointment information on screen
            Console.WriteLine("Summary: " + loadedAppointment.Summary);
            Console.WriteLine("Location: " + loadedAppointment.Location);
            Console.WriteLine("Description: " + loadedAppointment.Description);
            Console.WriteLine("Start date: " + loadedAppointment.StartDate);
            Console.WriteLine("End date: " + loadedAppointment.EndDate);
            Console.WriteLine("Organizer: " + appointment.Organizer);
            Console.WriteLine("Attendees: " + appointment.Attendees);
            Console.WriteLine(Environment.NewLine + "Appointment loaded successfully from " + dstEmail);
            // ExEnd:LoadAppointment
        }
    }
}
