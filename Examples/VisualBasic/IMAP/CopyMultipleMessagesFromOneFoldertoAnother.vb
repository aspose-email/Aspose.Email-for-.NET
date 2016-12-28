Imports Aspose.Email.Imap
Imports Aspose.Email.Mail

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CopyMultipleMessagesFromOneFoldertoAnother
        Public Shared Sub Run()
            'ExStart:CopyMultipleMessagesFromOneFoldertoAnother
            Using client As New ImapClient("exchange.aspose.com", "username", "password")
                ' Create the destination folder
                Dim folderName As String = "EMAILNET-35242"
                If Not client.ExistFolder(folderName) Then
                    client.CreateFolder(folderName)
                End If
                Try
                    ' Append a couple of messages to the server
                    Dim message1 As New MailMessage("asposeemail.test3@aspose.com", "asposeemail.test3@aspose.com", "EMAILNET-35242 - " + Guid.NewGuid().ToString(), "EMAILNET-35242 Improvement of copy method.Add ability to 'copy' multiple messages per invocation.")
                    Dim uniqueId1 As String = client.AppendMessage(message1)

                    Dim message2 As New MailMessage("asposeemail.test3@aspose.com", "asposeemail.test3@aspose.com", "EMAILNET-35242 - " + Guid.NewGuid().ToString(), "EMAILNET-35242 Improvement of copy method.Add ability to 'copy' multiple messages per invocation.")
                    Dim uniqueId2 As String = client.AppendMessage(message2)

                    ' Verify that the messages have been added to the mailbox
                    client.SelectFolder(ImapFolderInfo.InBox)
                    Dim msgsColl As ImapMessageInfoCollection = client.ListMessages()
                    For Each msgInfo As ImapMessageInfo In msgsColl
                        Console.WriteLine(msgInfo.Subject)
                    Next

                    ' Copy multiple messages using the CopyMessages command and Verify that messages are copied to the destination folder
                    client.CopyMessages(New String() {uniqueId1, uniqueId2}, folderName, True)
                    client.SelectFolder(folderName)
                    msgsColl = client.ListMessages()
                    For Each msgInfo As ImapMessageInfo In msgsColl
                        Console.WriteLine(msgInfo.Subject)
                    Next
                Finally
                    Try
                        client.DeleteFolder(folderName)
                    Catch
                    End Try
                End Try
            End Using
            'ExEnd:CopyMultipleMessagesFromOneFoldertoAnother
        End Sub
    End Class
End Namespace