Imports Aspose.Email.Exchange
Imports Aspose.Email.Outlook
Imports Aspose.Email.Mail

Public Class SendCalendarInvitation
    Public Shared Sub Run()
        'ExStart: SendCalendarInvitation
        ' <summary>
        ' This exmpale shows how an Exchange user can share his/her calendar with someone using the EWS client of the API. 
        ' Available from Aspose.Email for .NET 6.4.0 onwards
        ' </summary>
        Using client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")
            ' delegate calendar access permission
            Dim delegateUser As New ExchangeDelegateUser("sharingfrom@domain.com", ExchangeDelegateFolderPermissionLevel.NotSpecified)
            delegateUser.FolderPermissions.CalendarFolderPermissionLevel = ExchangeDelegateFolderPermissionLevel.Reviewer
            client.DelegateAccess(delegateUser, "sharingfrom@domain.com")

            ' create invitation
            Dim mapiMessage As MapiMessage = client.CreateCalendarSharingInvitationMessage("sharingfrom@domain.com")

            ' send invitation
            Dim messageInterpretor As MailMessageInterpretor = MailMessageInterpretorFactory.Instance.GetIntepretor(mapiMessage.MessageClass)
            Dim mailMessage As MailMessage = messageInterpretor.InterpretAsTnef(mapiMessage)
            client.Send(mailMessage)
        End Using
        'ExEnd: SendCalendarInvitation
    End Sub

End Class
