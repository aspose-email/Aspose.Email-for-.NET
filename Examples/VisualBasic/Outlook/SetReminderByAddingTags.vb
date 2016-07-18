Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class SetReminderByAddingTags
        Public Shared Sub Run()
            ' ExStart:SetReminderByAddingTags
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim location As String = "Meeting Location: Room 5"
            Dim startDate As New DateTime(1997, 3, 18, 18, 30, 0), endDate As New DateTime(1997, 3, 18, 19, 30, 0)
            Dim organizer As New MailAddress("aaa@amail.com", "Organizer")
            Dim attendees As New MailAddressCollection()
            attendees.Add(New MailAddress("bbb@bmail.com", "First attendee"))

            Dim target As New Appointment(location, startDate, endDate, organizer, attendees)

            ' Audio alarm that will sound at a precise time and
            ' Repeat 4 more times at 15 minute intervals:
            Dim audioReminder As New AppointmentReminder()
            audioReminder.Trigger = New ReminderTrigger(New DateTime(1997, 3, 17, 13, 30, 0, DateTimeKind.Utc))
            audioReminder.Repeat = 4
            audioReminder.Duration = New ReminderDuration(New TimeSpan(0, 15, 0))
            audioReminder.Action = ReminderAction.Audio
            Dim attach As New ReminderAttachment(New Uri("ftp://Host.com/pub/sounds/bell-01.aud"))
            audioReminder.Attachments.Add(attach)
            target.Reminders.Add(audioReminder)

            Dim strAudioReminder As String = vbCr & vbLf & "  BEGIN:VALARM" & vbCr & vbLf &
            "            ACTION:AUDIO" & vbCr & vbLf &
            "            REPEAT:4" & vbCr & vbLf &
            "            DURATION:PT15M" & vbCr & vbLf &
            "            TRIGGERVALUE=DATE-TIME:19970317T133000Z" & vbCr & vbLf &
            "            ATTACH:ftp:// Host.com/pub/sounds/bell-01.aud" & vbCr & vbLf &
            "            END:VALARM"

            ' Display alarm that will trigger 30 minutes before the
            ' Scheduled start of the event it is
            ' Associated with and will repeat 2 more times at 15 minute intervals:
            Dim displayReminder As New AppointmentReminder()
            Dim dur As New ReminderDuration(New TimeSpan(0, -30, 0))
            displayReminder.Trigger = New ReminderTrigger(dur, ReminderRelated.Start)
            displayReminder.Repeat = 2
            displayReminder.Duration = New ReminderDuration(New TimeSpan(0, 15, 0))
            displayReminder.Action = ReminderAction.Display
            displayReminder.Description = "Breakfast meeting with executive team at 8:30 AM EST"
            target.Reminders.Add(displayReminder)

            Dim strDisplayReminder As String = vbCr & vbLf & "            BEGIN:VALARM" & vbCr & vbLf &
                "            ACTION:DISPLAY" & vbCr & vbLf & "            REPEAT:2" & vbCr & vbLf &
                "            DURATION:PT15M" & vbCr & vbLf &
                "            DESCRIPTION:Breakfast meeting with executive team at 8:30 AM EST" & vbCr & vbLf &
                "            TRIGGERRELATED=START:-PT30M" & vbCr & vbLf & "            END:VALARM"

            ' Email alarm that will trigger 2 days before the
            ' Scheduled due date/time. It does not
            ' Repeat. The email has a subject, body and attachment link.
            Dim emailReminder As New AppointmentReminder()
            Dim dur1 As New ReminderDuration(New TimeSpan(-2, 0, 0, 0))
            emailReminder.Trigger = New ReminderTrigger(dur1, ReminderRelated.Start)
            Dim attendee As New ReminderAttendee("john_doe@host.com")
            emailReminder.Attendees.Add(attendee)
            emailReminder.Action = ReminderAction.Email
            emailReminder.Summary = "REMINDER: SEND AGENDA FOR WEEKLY STAFF MEETING"
            emailReminder.Description = "A draft agenda needs to be sent out to the attendees to the weekly managers meeting (MGR-LIST). Attached is a pointer the document template for the agenda file."
            Dim attach1 As New ReminderAttachment(New Uri("http://Host.com/templates/agenda.doc"))
            emailReminder.Attachments.Add(attach1)
            target.Reminders.Add(emailReminder)

            Dim strEmailReminder As String = vbCr & vbLf & "            BEGIN:VALARM" & vbCr & vbLf &
                "            ACTION:EMAIL" & vbCr & vbLf &
                "            DESCRIPTION:A draft agenda needs to be sent out to the attendees to the weekly managers meeting (MGR-LIST). Attached is a pointer the document template for the agenda file." & vbCr & vbLf &
                "            SUMMARY:REMINDER: SEND AGENDA FOR WEEKLY STAFF MEETING" & vbCr & vbLf & "            TRIGGERRELATED=START:-P2D" & vbCr & vbLf & "            ATTENDEE:mailto:john_doe@host.com" & vbCr & vbLf & "            ATTACH:http:// Host.com/templates/agenda.doc" & vbCr & vbLf & "            END:VALARM"

            ' Procedural alarm that will trigger at a precise date/time
            ' And will repeat 23 more times at one hour intervals. The alarm will
            ' Invoke a procedure file.
            Dim procReminder As New AppointmentReminder()
            procReminder.Trigger = New ReminderTrigger(New DateTime(1998, 1, 1, 5, 0, 0, _
                DateTimeKind.Utc))
            procReminder.Repeat = 23
            procReminder.Duration = New ReminderDuration(New TimeSpan(1, 0, 0))
            procReminder.Action = ReminderAction.Procedure
            Dim attach2 As New ReminderAttachment(New Uri("ftp://Host.com/novo-procs/felizano.exe"))
            procReminder.Attachments.Add(attach2)
            target.Reminders.Add(procReminder)

            Dim strProcReminder As String = vbCr & vbLf & "            BEGIN:VALARM" & vbCr & vbLf &
                "            ACTION:PROCEDURE" & vbCr & vbLf & "            REPEAT:23" & vbCr & vbLf &
                "            DURATION:PT1H" & vbCr & vbLf &
                "            TRIGGERVALUE=DATE-TIME:19980101T050000Z" & vbCr & vbLf &
                "            ATTACH:ftp:// Host.com/novo-procs/felizano.exe" & vbCr & vbLf &
                "            END:VALARM"

            target.Save(dataDir & Convert.ToString("savedFile_out.ics"))
            'ExEnd:SetReminderByAddingTags
        End Sub
    End Class
End Namespace
