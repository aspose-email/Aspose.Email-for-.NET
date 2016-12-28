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

    Public Class MeetingRequests
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SMTP()
            Dim dstEmail As String = dataDir & Convert.ToString("outputAttachments.msg")

            'Create an instance of the MailMessage class
            Dim msg As New MailMessage()

            'set the sender
            msg.From = "newcustomeronnet@gmail.com"

            'Set the recipient, who will receive the meeting request.
            'Basically, the recipient is the same as the meeting attendees.
            msg.[To] = "person1@domain.com, person2@domain.com, person3@domain.com, asposetest123@gmail.com"

            'create Appointment instance
            'location
            'start time
            'end time
            'organizer
            'attendee
            Dim app As New Appointment("Room 112", New DateTime(2015, 7, 17, 13, 0, 0), New DateTime(2015, 7, 17, 14, 0, 0), msg.From, msg.[To])

            app.Summary = "Release Meetting"
            app.Description = "Discuss for the next release"

            'add appointment to the message
            msg.AddAlternateView(app.RequestApointment())

            'Create an instance of SmtpClient class
            Dim client As SmtpClient = GetSmtpClient()

            Try
                'Client.Send will send this message
                client.Send(msg)
                'Show Message Sent�E only if message sent successfully
                Console.WriteLine("Message sent")

            Catch ex As Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Console.WriteLine(Environment.NewLine + "Meeting request send successfully.")
        End Sub

        Private Shared Function GetSmtpClient() As SmtpClient
            Dim client As New SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password")
            client.SecurityOptions = SecurityOptions.Auto

            Return client
        End Function
    End Class
End Namespace