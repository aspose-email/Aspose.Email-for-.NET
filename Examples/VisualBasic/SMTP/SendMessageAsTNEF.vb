Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.SMTP
    Class SendMessageAsTNEF
        Public Shared Sub Run()
            Try
                ' ExStart:SendMessageAsTNEF
                Dim emlFileName = RunExamples.GetDataDir_Email() + "Message.eml"
                ' A TNEF Email
                ' Load from eml
                Dim eml1 As MailMessage = MailMessage.Load(emlFileName, New EmlLoadOptions())
                eml1.From = "somename@gmail.com"
                eml1.[To].Clear()
                eml1.[To].Add(New MailAddress("first.last@test.com"))
                eml1.Subject = "With PreserveTnef flag during loading"
                eml1.[Date] = DateTime.Now
                Dim client As New SmtpClient("smtp.gmail.com", 587, "somename", "password")
                client.SecurityOptions = SecurityOptions.Auto
                client.UseTnef = True
                ' Use this flag to send as TNEF
                ' ExEnd:SendMessageAsTNEF
                client.Send(eml1)
            Catch ex As Exception
                Console.Write(ex.Message)
            End Try
        End Sub
    End Class
End Namespace