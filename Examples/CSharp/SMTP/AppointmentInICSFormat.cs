using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using Aspose.Email.Pop3;
using Aspose.Email;
using Aspose.Email.Mime;
using Aspose.Email.Imap;

namespace Aspose.Email.Examples.CSharp.SMTP
{
    class AppointmentInICSFormat
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SMTP();
            string dstEmail = dataDir + "test.ics";

            // 1. 
            // Create and save an Appointment to disk.

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

            // Display Status.
            System.Console.WriteLine("Appointment created and saved to disk successfully.");

            // 2.
            // Load an Appointment just created and saved to disk and display its details.

            // Load the appointment in ICS format
            Appointment loadedAppointment = Appointment.Load(dstEmail);

            // Display Status.
            System.Console.WriteLine(Environment.NewLine + "Loaded Appointment details are as follows:");

            // Display the appointment information on screen
            Console.WriteLine("Summary: " + loadedAppointment.Summary);
            Console.WriteLine("Location: " + loadedAppointment.Location);
            Console.WriteLine("Description: " + loadedAppointment.Description);
            Console.WriteLine("Start date: " + loadedAppointment.StartDate);
            Console.WriteLine("End date: " + loadedAppointment.EndDate);
            Console.WriteLine("Organizer: " + appointment.Organizer.ToString());
            Console.WriteLine("Attendees: " + appointment.Attendees.ToString());

            Console.WriteLine(Environment.NewLine + "Appointment loaded successfully from " + dstEmail);
        }
    }
}
