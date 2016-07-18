Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreatAndSaveCalendaritems
        Public Shared Sub Run()
            'ExStart:CreatAndSaveCalendaritems
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Create the appointment
            Dim calendar As New MapiCalendar("LAKE ARGYLE WA 6743", "Appointment", "This is a very important meeting :)", New DateTime(2012, 10, 2, 13, 0, 0), New DateTime(2012, 10, 2, 14, 0, 0))

            calendar.Save(dataDir & Convert.ToString("CalendarItem_out.ics"), AppointmentSaveFormat.Ics)
            'ExEnd:CreatAndSaveCalendaritems
        End Sub
    End Class
End Namespace
