Imports System.Collections.Generic
Imports System.Threading
Imports Aspose.Email.Imap
Imports Aspose.Email.Mail

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class SendIMAPasynchronousEmail
        Public Shared Sub Run()
            Try

                ' ExStart:SendIMAPasynchronousEmail
                ' Create an imapclient with host, user and password
                Dim client As New ImapClient()
                client.Host = "domain.com"
                client.Username = "username"
                client.Password = "password"
                client.SelectFolder("InBox")

                Dim messages As ImapMessageInfoCollection = client.ListMessages()
                Dim res1 As IAsyncResult = client.BeginFetchMessage(messages(0).UniqueId)
                Dim res2 As IAsyncResult = client.BeginFetchMessage(messages(1).UniqueId)
                Dim msg1 As MailMessage = client.EndFetchMessage(res1)
                Dim msg2 As MailMessage = client.EndFetchMessage(res2)

                ' ExEnd:SendIMAPasynchronousEmail

                ' ExStart:SendIMAPasynchronousEmailThreadPool
                Dim List As New List(Of MailMessage)()
                ThreadPool.QueueUserWorkItem(Sub(o As Object)
                                                 client.SelectFolder("folderName")
                                                 Dim messageInfoCol As ImapMessageInfoCollection = client.ListMessages()
                                                 For Each messageInfo As ImapMessageInfo In messageInfoCol
                                                     List.Add(client.FetchMessage(messageInfo.UniqueId))
                                                 Next

                                             End Sub)
                ' ExEnd:SendIMAPasynchronousEmailThreadPool

                ' ExStart:SendIMAPasynchronousEmailThreadPool1
                Dim List1 As New List(Of MailMessage)()
                ThreadPool.QueueUserWorkItem(Sub(o As Object)
                                                 Using connection As IDisposable = client.CreateConnection()
                                                     client.SelectFolder("FolderName")
                                                     Dim messageInfoCol As ImapMessageInfoCollection = client.ListMessages()
                                                     For Each messageInfo As ImapMessageInfo In messageInfoCol
                                                         List1.Add(client.FetchMessage(messageInfo.UniqueId))
                                                     Next
                                                     ' ExEnd:SendIMAPasynchronousEmailThreadPool1
                                                 End Using
                                             End Sub)
            Catch ex As Exception
                Console.Write(ex.Message)
                Throw
            End Try
        End Sub
    End Class
End Namespace