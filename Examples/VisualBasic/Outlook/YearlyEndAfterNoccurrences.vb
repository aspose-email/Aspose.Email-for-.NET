Imports Aspose.Email.Outlook

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class YearlyEndAfterNoccurrences
        Public Shared Sub Run()
            ' ExStart:YearlyEndAfterNoccurrences
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim localZone As TimeZone = TimeZone.CurrentTimeZone
            Dim ts As TimeSpan = localZone.GetUtcOffset(DateTime.Now)
            Dim StartDate As New DateTime(2015, 7, 1)
            StartDate = StartDate.Add(ts)

            Dim DueDate As New DateTime(2015, 7, 1)
            Dim endByDate As New DateTime(2020, 12, 31)
            DueDate = DueDate.Add(ts)
            endByDate = endByDate.Add(ts)

            Dim task As New MapiTask("This is test task", "Sample Body", StartDate, DueDate)
            task.State = MapiTaskState.NotAssigned

            ' Set the Yearly recurrence
            Dim rec = New MapiCalendarMonthlyRecurrencePattern() With { _
                .Day = 15, _
                .Period = 12, _
                .PatternType = MapiCalendarRecurrencePatternType.Month, _
                .EndType = MapiCalendarRecurrenceEndType.EndAfterNOccurrences, _
                .OccurrenceCount = 3 _
            }

            If rec.OccurrenceCount = 0 Then
                rec.OccurrenceCount = 1
            End If

            task.Recurrence = rec
            task.Save(dataDir & Convert.ToString("Yearly_out.msg"), TaskSaveFormat.Msg)
            ' ExSEnd:YearlyEndAfterNoccurrences
        End Sub

    End Class
End Namespace
