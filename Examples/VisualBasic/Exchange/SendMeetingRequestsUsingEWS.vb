Imports System.Net
Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail
Imports Aspose.Email.Mail.Calendaring

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
    Class SendMeetingRequestsUsingEWS
        Public Shared Sub Run()
            ' ExStart:SendMeetingRequestsUsingEWS
            Try
                ' Create instance of IEWSClient class by giving credentials
                Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

                ' create the meeting request
                Dim app As New Appointment("meeting request", DateTime.Now.AddHours(1), DateTime.Now.AddHours(1.5), "administrator@test.com", "bob@test.com")
                app.Summary = "meeting request summary"
                app.Description = "description"

                Dim pattern As Aspose.Email.Recurrences.RecurrencePattern = New Recurrences.DailyRecurrencePattern(DateTime.Now.AddDays(5))
                app.Recurrence = pattern

                ' create the message and set the meeting request
                Dim msg As New MailMessage()
                msg.From = "administrator@test.com"
                msg.[To] = "bob@test.com"
                msg.IsBodyHtml = True
                msg.HtmlBody = "<h3>HTML Heading</h3><p>Email Message detail</p>"
                msg.Subject = "meeting request"
                msg.AddAlternateView(app.RequestApointment(0))

                ' send the appointment
                client.Send(msg)
                Console.WriteLine("Appointment request sent")
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
            ' ExEnd:SendMeetingRequestsUsingEWS
        End Sub
    End Class
End Namespace