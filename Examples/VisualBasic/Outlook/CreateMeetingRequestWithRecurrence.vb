Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreateMeetingRequestWithRecurrence

        Public Shared Sub Run()

            Try
                ' ExStart:CreateMeetingRequestWithRecurrence
                Dim szUniqueId As [String]

                ' Create a mail message
                Dim msg1 As New MailMessage()
                msg1.[To].Add("to@domain.com")
                msg1.From = New MailAddress("from@gmail.com")

                ' Create appointment object
                Dim agendaAppointment As Appointment = Nothing

                ' Fill appointment object
                Dim StartDate As System.DateTime = New DateTime(2013, 12, 1, 17, 0, 0)
                Dim EndDate As System.DateTime = New DateTime(2013, 12, 31, 17, 30, 0)
                agendaAppointment = New Appointment("same place", StartDate, EndDate, msg1.From, msg1.[To])

                ' Create unique id as it will help to access this appointment later
                szUniqueId = Guid.NewGuid().ToString()
                agendaAppointment.UniqueId = szUniqueId
                agendaAppointment.Description = "----------------"

                ' Create a weekly reccurence pattern object
                Dim pattern2 As New Aspose.Email.Mail.Calendaring.WeeklyRecurrencePattern(14)

                ' Set weekly pattern properties like days: Mon, Tue and Thu
                pattern2.StartDays = New Aspose.Email.Mail.Calendaring.CalendarDay(2) {}
                pattern2.StartDays(0) = Aspose.Email.Mail.Calendaring.CalendarDay.Monday
                pattern2.StartDays(1) = Aspose.Email.Mail.Calendaring.CalendarDay.Tuesday
                pattern2.StartDays(2) = Aspose.Email.Mail.Calendaring.CalendarDay.Thursday
                pattern2.Interval = 1

                ' Set recurrence pattern for the appointment
                agendaAppointment.RecurrencePattern = pattern2

                'Attach this appointment with mail
                msg1.AlternateViews.Add(agendaAppointment.RequestApointment())

                ' Create SmtpCleint
                Dim client As New SmtpClient("smtp.gmail.com", 587, "your.email@gmail.com", "your.password")
                client.SecurityOptions = SecurityOptions.Auto

                ' Send mail with appointment request
                client.Send(msg1)

                ' Return unique id for later usage
                ' return szUniqueId;
                ' ExEnd:SendMailUsingDNS
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Sub
    End Class
End Namespace