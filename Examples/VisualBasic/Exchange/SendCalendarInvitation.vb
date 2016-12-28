Imports Aspose.Email.Exchange
Imports Aspose.Email.Mail
Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
    Class SendCalendarInvitation
        Public Shared Sub Run()
            Try
                ' ExStart: SendCalendarInvitation
                Using client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")
                    ' delegate calendar access permission
                    Dim delegateUser As New ExchangeDelegateUser("sharingfrom@domain.com", ExchangeDelegateFolderPermissionLevel.NotSpecified)
                    delegateUser.FolderPermissions.CalendarFolderPermissionLevel = ExchangeDelegateFolderPermissionLevel.Reviewer
                    client.DelegateAccess(delegateUser, "sharingfrom@domain.com")

                    ' Create invitation
                    Dim mapiMessage As MapiMessage = client.CreateCalendarSharingInvitationMessage("sharingfrom@domain.com")
                    Dim options As New MailConversionOptions()
                    options.ConvertAsTnef = True
                    Dim mail As MailMessage = mapiMessage.ToMailMessage(options)
                    client.Send(mail)
                    ' ExEnd: SendCalendarInvitation
                End Using
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Sub
    End Class
End Namespace