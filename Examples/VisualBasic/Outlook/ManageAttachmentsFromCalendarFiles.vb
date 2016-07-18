Imports System.IO
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class ManageAttachmentsFromCalendarFiles
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim files As String() = New String(2) {}
            files(0) = dataDir & Convert.ToString("attachment_1.doc")
            files(1) = dataDir & Convert.ToString("download.png")
            files(2) = dataDir & Convert.ToString("Desert.jpg")

            Dim app1 As New Appointment("Home", DateTime.Now.AddHours(1), DateTime.Now.AddHours(1), "organizer@domain.com", "attendee@gmail.com")
            For Each file__1 As String In files
                Using ms As New MemoryStream(File.ReadAllBytes(file__1))
                    app1.Attachments.Add(New Attachment(ms, Path.GetFileName(file__1)))
                End Using
            Next

            app1.Save(dataDir & Convert.ToString("appWithAttachments_out.ics"), AppointmentSaveFormat.Ics)

            Dim app2 As Appointment = Appointment.Load(dataDir & Convert.ToString("appWithAttachments_out.ics"))
            Console.WriteLine(app2.Attachments.Count)
            For Each att As Attachment In app2.Attachments
                Console.WriteLine(att.Name)
            Next

        End Sub
    End Class
End Namespace
