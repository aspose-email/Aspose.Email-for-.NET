Imports Aspose.Email.Exchange

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Exchange
    Class ListMessagesByExtendedProperties
        Public Shared Sub Run()
            Try
                ' ExStart:ListMessagesByExtendedProperties
                Using client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")
                    Dim messageInfoCol As ExchangeMessageInfoCollection = client.ListMessages(client.MailboxInfo.InboxUri, ExchangeListMessagesOptions.FetchAttachmentInformation, New String() {"blablaprop"})
                    ' ExEnd:ListMessagesByExtendedProperties
                End Using
            Catch ex As Exception
                Console.WriteLine(ex.Message + "Please check your credentials")
            Finally
            End Try

        End Sub
    End Class
End Namespace
