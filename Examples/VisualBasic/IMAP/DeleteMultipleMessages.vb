
Imports System.Collections.Generic
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
    Class DeleteMultipleMessages
        Public Shared Sub Run()
            'ExStart:DeleteMultipleMessages
            Using client As New ImapClient("exchange.aspose.com", "username", "password")
                Try
                    Console.WriteLine(client.UidPlusSupported.ToString())
                    ' Append some test messages
                    client.SelectFolder(ImapFolderInfo.InBox)
                    Dim uidList As New List(Of String)()
                    Const messageNumber As Integer = 5
                    For i As Integer = 0 To messageNumber - 1
                        Dim message As New MailMessage("from@Aspose.com", "to@Aspose.com", "EMAILNET-35226 - " + Guid.NewGuid().ToString(), "EMAILNET-35226 Add ability in ImapClient to delete messages and change flags for set of messages")
                        Dim uid As String = client.AppendMessage(message)
                        uidList.Add(uid)
                    Next

                    ' Now verify that all the messages have been appended to the mailbox
                    Dim messageInfoCol As ImapMessageInfoCollection = Nothing
                    messageInfoCol = client.ListMessages()
                    Console.WriteLine(messageInfoCol.Count)

                    ' Bulk Delete Messages and  Verify that the messages are deleted
                    client.DeleteMessages(uidList, True)
                    client.CommitDeletes()
                    messageInfoCol = Nothing
                    messageInfoCol = client.ListMessages()
                    Console.WriteLine(messageInfoCol.Count)

                Finally
                End Try
            End Using
            'ExStart:DeleteMultipleMessages
        End Sub
    End Class
End Namespace