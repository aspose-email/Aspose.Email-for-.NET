Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class AddRecurrenceToMapiTask
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim startDate As New DateTime(2015, 4, 30, 10, 0, 0)
            Dim task As New MapiTask("abc", "def", startDate, startDate.AddHours(1))
            task.State = MapiTaskState.NotAssigned

            ' Set the weekly recurrence
            Dim rec = New MapiCalendarDailyRecurrencePattern() With { _
                .PatternType = MapiCalendarRecurrencePatternType.Day, _
                .Period = 1, _
                .WeekStartDay = DayOfWeek.Sunday, _
                .EndType = MapiCalendarRecurrenceEndType.NeverEnd, _
                .OccurrenceCount = 0 _
            }
            task.Recurrence = rec
            task.Save(dataDir & Convert.ToString("AsposeDaily_out.msg"), TaskSaveFormat.Msg)

            ' Set the weekly recurrence
            ' Set the weekly recurrence
            Dim rec1 = New MapiCalendarWeeklyRecurrencePattern() With { _
                .PatternType = MapiCalendarRecurrencePatternType.Week, _
                .Period = 1, _
                .DayOfWeek = MapiCalendarDayOfWeek.Wednesday, _
                .EndType = MapiCalendarRecurrenceEndType.NeverEnd, _
                .OccurrenceCount = 0 _
            }
            task.Recurrence = rec1
            task.Save(dataDir & Convert.ToString("AsposeWeekly_out.msg"), TaskSaveFormat.Msg)

            ' Set the monthly recurrence
            Dim recMonthly = New MapiCalendarMonthlyRecurrencePattern() With { _
                .PatternType = MapiCalendarRecurrencePatternType.Month, _
                .Period = 1, _
                .EndType = MapiCalendarRecurrenceEndType.NeverEnd, _
                .Day = 30, _
                .OccurrenceCount = 0, _
                .WeekStartDay = DayOfWeek.Sunday _
            }
            task.Recurrence = recMonthly
            task.Save(dataDir & Convert.ToString("AsposeMonthly_out.msg"), TaskSaveFormat.Msg)

            ' Set the yearly recurrence
            Dim recYearly = New MapiCalendarMonthlyRecurrencePattern() With { _
                .PatternType = MapiCalendarRecurrencePatternType.Month, _
                .EndType = MapiCalendarRecurrenceEndType.NeverEnd, _
                .OccurrenceCount = 10, _
                .Period = 12 _
            }
            task.Recurrence = recYearly
            task.Save(dataDir & Convert.ToString("AsposeYearly_out.msg"), TaskSaveFormat.Msg)
        End Sub
    End Class
End Namespace
