Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.SMTP
    Class SendEmailWithAlternateText
        Public Shared Sub Run()
            ' ExStart:SendEmailWithAlternateText
            'Create an instance of the MailMessage class
            Dim message As New MailMessage()

            ' Set From field, To field and Plain text body
            message.From = "sender@sender.com"
            message.[To].Add("receiver@receiver.com")
            message.Body = "This is Plain Text Body"

            ' Create an instance of the SmtpClient class
            Dim client As New SmtpClient()

            ' And Specify your mailing host server, Username, Password and Port
            client.Host = "smtp.server.com"
            client.Username = "Username"
            client.Password = "Password"
            client.Port = 25

            Try
                ' Client.Send will send this message
                client.Send(message)
                Console.WriteLine("Message sent")
            Catch ex As Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Console.WriteLine("Press enter to quit")
            Console.Read()
        End Sub
        ' ExEnd:SendEmailWithAlternateText
    End Class
End Namespace