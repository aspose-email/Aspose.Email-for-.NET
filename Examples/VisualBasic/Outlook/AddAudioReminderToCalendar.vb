Imports Aspose.Email.Outlook
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class AddAudioReminderToCalendar
        Public Shared Sub Run()
            ' ExStart:AddAudioReminderToCalendar
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim app As New Appointment("Home", DateTime.Now.AddHours(1), DateTime.Now.AddHours(1), "organizer@domain.com", "attendee@gmail.com")

            Dim msg As New MailMessage()
            msg.AddAlternateView(app.RequestApointment())
            Dim mapi As MapiMessage = MapiMessage.FromMailMessage(msg)
            Dim calendar As MapiCalendar = DirectCast(mapi.ToMapiMessageItem(), MapiCalendar)

            ' Set calendar properties
            calendar.ReminderSet = True
            calendar.ReminderDelta = 58
            '58 min before start of event
            calendar.ReminderFileParameter = dataDir & Convert.ToString("Alarm01.wav")
            Dim savedFile As String = (dataDir & Convert.ToString("calendarWithAudioReminder_out.ics"))
            calendar.Save(savedFile, AppointmentSaveFormat.Ics)
            ' ExEnd:AddAudioReminderToCalendar
        End Sub
    End Class
End Namespace
