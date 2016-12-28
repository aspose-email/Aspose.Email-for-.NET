Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Aspose.Email.Outlook.Pst
Imports Aspose.Email
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
    Class WeeklyEndAfterNoccurrences
        Public Shared Sub Run()
            ' ExStart:WeeklyEndAfterNoccurrences
            ' The path to the File directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim localZone As TimeZone = TimeZone.CurrentTimeZone
            Dim ts As TimeSpan = localZone.GetUtcOffset(DateTime.Now)
            Dim StartDate As New DateTime(2015, 7, 16)
            StartDate = StartDate.Add(ts)

            Dim DueDate As New DateTime(2015, 7, 16)
            Dim endByDate As New DateTime(2015, 9, 1)
            DueDate = DueDate.Add(ts)
            endByDate = endByDate.Add(ts)

            Dim task As New MapiTask("This is test task", "Sample Body", StartDate, DueDate)
            task.State = MapiTaskState.NotAssigned

            ' Set the weekly recurrence
            Dim rec = New MapiCalendarWeeklyRecurrencePattern() With { _
                 .EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences, _
                .PatternType = MapiCalendarRecurrencePatternType.Week, _
                .Period = 1, _
                .WeekStartDay = DayOfWeek.Sunday, _
                .DayOfWeek = MapiCalendarDayOfWeek.Friday, _
                .OccurrenceCount = GetOccurrenceCount(StartDate, endByDate, "FREQ=WEEKLY;BYDAY=FR") _
            }

            If rec.OccurrenceCount = 0 Then
                rec.OccurrenceCount = 1
            End If

            task.Recurrence = rec
            task.Save(dataDir & "Weekly_out.msg", TaskSaveFormat.Msg)
            ' ExEnd:WeeklyEndAfterNoccurrences
        End Sub
        Private Shared Function GetOccurrenceCount(start As DateTime, endBy As DateTime, rrule As String) As UInteger
            ' ExStart:GetOccurrenceCount
            Dim pattern As New CalendarRecurrence(String.Format("DTSTART:{0}" & vbCr & vbLf & "RRULE:{1}", start.ToString("yyyyMMdd"), rrule))
            Dim dates As DateCollection = pattern.GenerateOccurrences(start, endBy)
            Return CUInt(dates.Count)
            ' ExEnd:GetOccurrenceCount
        End Function
    End Class
End Namespace
