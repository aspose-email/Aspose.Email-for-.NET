Imports System.IO
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
    Class DownloadMessagesFromExchangeServerFoldersRecursively
        ' ExStart:DownloadMessagesFromExchangeServerFoldersRecursively
        Private Shared mailboxURI As String = "https://ex2010/ews/exchange.asmx" ' EWS
        Private Shared username As String = "administrator"
        Private Shared password As String = "pwd"
        Private Shared domain As String = "ex2010.local"

        Shared Sub Run()
            Try
                DownloadAllMessages()
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
            ' ExEnd:DownloadMessagesFromExchangeServerFoldersRecursively
        End Sub

        ' ExStart:DownloadAllMessages
        Private Shared Sub DownloadAllMessages()
            Try
                Dim rootFolder As String = domain & "-" & username
                Directory.CreateDirectory(rootFolder)
                Dim inboxFolder As String = rootFolder & "\Inbox"
                Directory.CreateDirectory(inboxFolder)

                Console.WriteLine("Downloading all messages from Inbox....")
                ' Create instance of EWSClient class by giving credentials
                Dim client As IEWSClient = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain")

                Dim mailboxInfo As ExchangeMailboxInfo = client.GetMailboxInfo()
                Console.WriteLine("Mailbox URI: " & mailboxInfo.MailboxUri)
                Dim rootUri As String = client.GetMailboxInfo().RootUri
                ' List all the folders from Exchange server
                Dim folderInfoCollection As ExchangeFolderInfoCollection = client.ListSubFolders(rootUri)
                For Each folderInfo As ExchangeFolderInfo In folderInfoCollection
                    ' Call the recursive method to read messages and get sub-folders
                    ListMessagesInFolder(client, folderInfo, rootFolder)
                Next folderInfo

                Console.WriteLine("All messages downloaded.")
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        End Sub

        ' Recursive method to get messages from folders and sub-folders
        Private Shared Sub ListMessagesInFolder(ByVal client As IEWSClient, ByVal folderInfo As ExchangeFolderInfo, ByVal rootFolder As String)
            ' Create the folder in disk (same name as on IMAP server)
            Dim currentFolder As String = rootFolder & "\" & folderInfo.DisplayName
            Directory.CreateDirectory(currentFolder)

            ' List messages
            Dim msgInfoColl As ExchangeMessageInfoCollection = client.ListMessages(folderInfo.Uri)
            Console.WriteLine("Listing messages....")
            Dim i As Integer = 0
            For Each msgInfo As ExchangeMessageInfo In msgInfoColl
                ' Get subject and other properties of the message
                Console.WriteLine("Subject: " & msgInfo.Subject)

                ' Get rid of characters like ? and :, which should not be included in a file name
                Dim fileName As String = msgInfo.Subject.Replace(":", " ").Replace("?", " ")

                ' Save the message in MSG format
                Dim msg As MailMessage = client.FetchMessage(msgInfo.UniqueUri)
                'msg.Save(currentFolder & "\" & fileName & "-" & i & ".msg", MailMessageSaveType.OutlookMessageFormat)
                i += 1
            Next msgInfo
            Console.WriteLine("============================" & Constants.vbLf)
            Try
                ' If this folder has sub-folders, call this method recursively to get messages
                Dim folderInfoCollection As ExchangeFolderInfoCollection = client.ListSubFolders(folderInfo.Uri)
                For Each subfolderInfo As ExchangeFolderInfo In folderInfoCollection
                    ListMessagesInFolder(client, subfolderInfo, currentFolder)
                Next subfolderInfo
            Catch e1 As Exception
            End Try
        End Sub
    End Class
    ' ExEnd:DownloadAllMessages
End Namespace
