Imports Aspose.Email.Outlook

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class GenerateRecurrenceFromRecurrenceRule
        Public Shared Sub Run()
            ' ExStart:GenerateRecurrenceFromRecurrenceRule
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim startDate As New DateTime(2015, 7, 16)
            Dim endDate As New DateTime(2015, 8, 1)
            Dim app1 As New MapiCalendar("test location", "test summary", "test description", startDate, endDate)

            app1.StartDate = startDate
            app1.EndDate = endDate

            Dim pattern As String = "DTSTART;TZID=Europe/London:20150831T080000" & vbCr & vbLf & "DTEND;TZID=Europe/London:20150831T083000" & vbCr & vbLf & "RRULE:FREQ=DAILY;INTERVAL=1;COUNT=7" & vbCr & vbLf & "EXDATE:20150831T070000Z,20150904T070000Z"
            app1.Recurrence.RecurrencePattern = MapiCalendarRecurrencePatternFactory.FromString(pattern)
            ' ExEnd:GenerateRecurrenceFromRecurrenceRule
        End Sub
    End Class
End Namespace
