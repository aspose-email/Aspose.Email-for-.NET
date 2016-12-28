Imports Aspose.Email.Outlook
Imports Aspose.Email.Recurrences

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class MonthlyEndAfterNoccurrences
        Public Shared Sub Run()
            ' ExStart:MonthlyEndAfterNoccurrences
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim localZone As TimeZone = TimeZone.CurrentTimeZone
            Dim ts As TimeSpan = localZone.GetUtcOffset(DateTime.Now)
            Dim StartDate As New DateTime(2015, 7, 16)
            StartDate = StartDate.Add(ts)

            Dim DueDate As New DateTime(2015, 7, 16)
            Dim endByDate As New DateTime(2015, 12, 31)
            DueDate = DueDate.Add(ts)
            endByDate = endByDate.Add(ts)

            Dim task As New MapiTask("This is test task", "Sample Body", StartDate, DueDate)
            task.State = MapiTaskState.NotAssigned

            ' Set the Monthly recurrence

            Dim rec = New MapiCalendarMonthlyRecurrencePattern() With { _
                .Day = 15, _
                .Period = 1, _
                .PatternType = MapiCalendarRecurrencePatternType.Month, _
                .EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences, _
                .OccurrenceCount = GetOccurrenceCount(StartDate, endByDate, "FREQ=MONTHLY;BYMONTHDAY=15;INTERVAL=1"), _
                .WeekStartDay = DayOfWeek.Monday _
            }

            If rec.OccurrenceCount = 0 Then
                rec.OccurrenceCount = 1
            End If
            task.Recurrence = rec

            task.Save(dataDir & "Monthly_out.msg", TaskSaveFormat.Msg)
            ' ExEnd:MonthlyEndAfterNoccurrences

        End Sub

        Private Shared Function GetOccurrenceCount(start As DateTime, endBy As DateTime, rrule As String) As UInteger
            ' ExStart:GetOccurrenceCountMonthlyEndAfterNoccurrences             
            Dim pattern As New CalendarRecurrence(String.Format("DTSTART:{0}" & vbCr & vbLf & "RRULE:{1}", start.ToString("yyyyMMdd"), rrule))
            Dim dates As DateCollection = pattern.GenerateOccurrences(start, endBy)
            Return CUInt(dates.Count)
            ' ExEnd:GetOccurrenceCountMonthlyEndAfterNoccurrences
        End Function
    End Class
End Namespace
