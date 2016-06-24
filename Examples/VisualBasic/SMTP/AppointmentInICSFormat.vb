Imports System.IO
Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook
Imports Aspose.Email.Pop3
Imports Aspose.Email
Imports Aspose.Email.Mime
Imports Aspose.Email.Imap
Imports System.Configuration
Imports System.Data
Imports Aspose.Email.Mail.Bounce
Imports Aspose.Email.Exchange
Imports Aspose.Email.Outlook.Pst

Namespace Aspose.Email.Examples.VisualBasic.Email.SMTP

    Public Class AppointmentInICSFormat
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim dstEmail As String = dataDir & Convert.ToString("test.ics")

            ' 1. 
            ' Create and save an Appointment to disk.

            ' Create and initialize an instance of the Appointment class
            ' Location
            ' Summary
            ' Description
            ' Start date
            ' End date
            ' Organizer
            Dim appointment__1 As New Appointment("Meeting Room 3 at Office Headquarters", "Monthly Meeting", "Please confirm your availability.", New DateTime(2015, 2, 8, 13, 0, 0), New DateTime(2015, 2, 8, 14, 0, 0), "from@domain.com", _
                "attendees@domain.com")
            ' Attendees
            ' Save the appointment to disk in ICS format
            appointment__1.Save(dstEmail, AppointmentSaveFormat.Ics)

            ' Display Status.
            System.Console.WriteLine("Appointment created and saved to disk successfully.")

            ' 2.
            ' Load an Appointment just created and saved to disk and display its details.

            ' Load the appointment in ICS format
            Dim loadedAppointment As Appointment = Appointment.Load(dstEmail)

            ' Display Status.
            System.Console.WriteLine(Environment.NewLine + "Loaded Appointment details are as follows:")

            ' Display the appointment information on screen
            Console.WriteLine("Summary: " + loadedAppointment.Summary)
            Console.WriteLine("Location: " + loadedAppointment.Location)
            Console.WriteLine("Description: " + loadedAppointment.Description)
            Console.WriteLine("Start date: " + loadedAppointment.StartDate)
            Console.WriteLine("End date: " + loadedAppointment.EndDate)
            Console.WriteLine("Organizer: " + appointment__1.Organizer.ToString())
            Console.WriteLine("Attendees: " + appointment__1.Attendees.ToString())

            Console.WriteLine(Convert.ToString(Environment.NewLine + "Appointment loaded successfully from ") & dstEmail)
        End Sub
    End Class
End Namespace