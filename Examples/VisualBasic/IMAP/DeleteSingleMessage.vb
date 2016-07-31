
Imports Aspose.Email.Imap
Imports Aspose.Email.Mail

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class DeleteSingleMessage
        Public Shared Sub Run()
            'ExStart:DeleteSingleMessage
            Using client As New ImapClient("exchange.aspose.com", "username", "password")
                Try
                    Console.WriteLine(client.UidPlusSupported.ToString())
                    ' Append some test messages
                    client.SelectFolder(ImapFolderInfo.InBox)
                    Dim message As New MailMessage("from@Aspose.com", "to@Aspose.com", "EMAILNET-35227 - " + Guid.NewGuid().ToString(), "EMAILNET-35227 Add ability in ImapClient to delete message")
                    Dim emailId As String = client.AppendMessage(message)

                    ' Now verify that all the messages have been appended to the mailbox
                    Dim messageInfoCol As ImapMessageInfoCollection = Nothing
                    messageInfoCol = client.ListMessages()
                    Console.WriteLine(messageInfoCol.Count)

                    ' Select the inbox folder and Delete message
                    client.SelectFolder(ImapFolderInfo.InBox)
                    client.DeleteMessage(emailId)
                    client.CommitDeletes()

                Finally
                End Try
            End Using
            'ExEnd:DeleteSingleMessage
        End Sub
    End Class
End Namespace