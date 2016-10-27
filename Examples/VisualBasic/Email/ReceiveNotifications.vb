Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email
    Class ReceiveNotifications
        Public Shared Sub Run()
            ' ExStart:ReceiveNotifications
            ' Create the message
            Dim msg As New MailMessage()
            msg.From = "sender@sender.com"
            msg.[To] = "receiver@receiver.com"
            msg.Subject = "the subject of the message"

            ' Set delivery notifications for success and failed messages
            msg.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess Or DeliveryNotificationOptions.OnFailure

            ' Add the MIME headers
            msg.Headers.Add_("Disposition-Notification-To", "sender@sender.com")
            msg.Headers.Add_("Disposition-Notification-To", "sender@sender.com")

            ' Send the message
            Try
                Dim client As New SmtpClient("host", "username", "password")
                client.Send(msg)
            Catch ex As Exception
                Console.WriteLine(ex.Message + "Please check your credentials")
            End Try
            ' ExEnd:ReceiveNotifications
        End Sub
    End Class
End Namespace