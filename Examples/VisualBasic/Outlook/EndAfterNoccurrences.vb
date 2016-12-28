Imports Aspose.Email.Outlook
Imports Aspose.Email.Recurrences

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class EndAfterNoccurrences
        Public Shared Sub Run()

            ' ExStart:EndAfterNoccurrences
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim localZone As TimeZone = TimeZone.CurrentTimeZone
            Dim timeSpan As TimeSpan = localZone.GetUtcOffset(DateTime.Now)
            Dim StartDate As New DateTime(2015, 7, 16)
            StartDate = StartDate.Add(timeSpan)

            Dim DueDate As New DateTime(2015, 7, 16)
            Dim endByDate As New DateTime(2015, 8, 1)
            DueDate = DueDate.Add(timeSpan)
            endByDate = endByDate.Add(timeSpan)

            Dim task As New MapiTask("This is test task", "Sample Body", StartDate, DueDate)
            task.State = MapiTaskState.NotAssigned

            ' Set the Daily recurrence
            Dim rec = New MapiCalendarDailyRecurrencePattern() With { _
                .PatternType = MapiCalendarRecurrencePatternType.Day, _
                .Period = 1, _
                .WeekStartDay = DayOfWeek.Sunday, _
                .EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences, _
                .OccurrenceCount = GetOccurrenceCount(StartDate, endByDate, "FREQ=DAILY") _
            }

            If rec.OccurrenceCount = 0 Then
                rec.OccurrenceCount = 1
            End If

            task.Recurrence = rec
            task.Save(dataDir & Convert.ToString("Daily_out.msg"), TaskSaveFormat.Msg)
            ' ExEnd:EndAfterNoccurrences
        End Sub
        Private Shared Function GetOccurrenceCount(start As DateTime, endBy As DateTime, rrule As String) As UInteger
            ' ExStart:GetOccurrenceCount
            Dim pattern As New CalendarRecurrence(String.Format("DTSTART:{0}" & vbCr & vbLf & "RRULE:{1}", start.ToString("yyyyMMdd"), rrule))
            Dim dates As DateCollection = pattern.GenerateOccurrences(start, endBy)
            Return CUInt(dates.Count)
            ' ExStart:GetOccurrenceCount
        End Function
    End Class
End Namespace
