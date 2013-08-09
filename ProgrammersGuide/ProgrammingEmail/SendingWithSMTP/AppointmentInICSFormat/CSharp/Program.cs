//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Email. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;

using Aspose.Email;
using Aspose.Email.Mail;

namespace AppointmentInICSFormat
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");
            Directory.CreateDirectory(dataDir);

            // 1. 
            // Create and save an Appointment to disk.

            // Create and initialize an instance of the Appointment class
            Appointment appointment = new Appointment(
                "Meeting Room 3 at Office Headquarters",// Location
                "Monthly Meeting",                      // Summary
                "Please confirm your availability.",    // Description
                new DateTime(2011, 2, 8, 13, 0, 0),     // Start date
                new DateTime(2011, 2, 8, 14, 0, 0),     // End date
                "from@domain.com",                      // Organizer
                "attendees@domain.com");                // Attendees

            // Save the appointment to disk in ICS format
            appointment.Save(dataDir + "test.ics", AppointmentSaveFormat.Ics);

            // Display Status.
            System.Console.WriteLine("Appointment created and saved to disk successfully.");

            // 2.
            // Load an Appointment just created and saved to disk and display its details.

            // Load the appointment in ICS format
            Appointment loadedAppointment = Appointment.Load(dataDir + "test.ics");

            // Display Status.
            System.Console.WriteLine("\n\nLoaded Appointment details are as follows:");

            // Display the appointment information on screen
            Console.WriteLine("Summary: " + loadedAppointment.Summary);
            Console.WriteLine("Location: " + loadedAppointment.Location);
            Console.WriteLine("Description: " + loadedAppointment.Description);
            Console.WriteLine("Start date: " + loadedAppointment.StartDate);
            Console.WriteLine("End date: " + loadedAppointment.EndDate);
            Console.WriteLine("Organizer: " + appointment.Organizer.ToString());
            Console.WriteLine("Attendees: " + appointment.Attendees.ToString());

            // Display Status.
            System.Console.WriteLine("\nAppointment loaded and information displayed successfully.");
        }
    }
}