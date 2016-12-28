Imports System.IO
Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class AddMapiCalendarToPST
        Public Shared Sub Run()
            ' The path to the File directory.
            ' ExStart:AddMapiCalendarToPST
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Create the appointment
            Dim appointment As New MapiCalendar("LAKE ARGYLE WA 6743", "Appointment", "This is a very important meeting :)", New DateTime(2012, 10, 2, 13, 0, 0), New DateTime(2012, 10, 2, 14, 0, 0))

            ' Create the meeting
            Dim attendees As New MapiRecipientCollection()
            attendees.Add("ReneeAJones@armyspy.com", "Renee A. Jones", MapiRecipientType.MAPI_TO)
            attendees.Add("SzllsyLiza@dayrep.com", "Szollosy Liza", MapiRecipientType.MAPI_TO)

            Dim meeting As New MapiCalendar("Meeting Room 3 at Office Headquarters", "Meeting", "Please confirm your availability.", New DateTime(2012, 10, 2, 13, 0, 0), New DateTime(2012, 10, 2, 14, 0, 0), "CharlieKhan@dayrep.com", _
                attendees)

            Dim checkfile As String = dataDir + "AddMapiCalendarToPST_out.pst"

            If File.Exists(checkfile) Then
                File.Delete(checkfile)
            Else
            End If

            Using pst As PersonalStorage = PersonalStorage.Create(dataDir & Convert.ToString("AddMapiCalendarToPST_out.pst"), FileFormatVersion.Unicode)
                Dim calendarFolder As FolderInfo = pst.CreatePredefinedFolder("Calendar", StandardIpmFolder.Appointments)
                calendarFolder.AddMapiMessageItem(appointment)
                calendarFolder.AddMapiMessageItem(meeting)
            End Using
            ' ExEnd:AddMapiCalendarToPST
        End Sub
    End Class
End Namespace
