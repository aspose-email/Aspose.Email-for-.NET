Imports Aspose.Email.Imap

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.IMAP
    Class IMAP4ExtendedListCommand
        Public Shared Sub Run()
            'ExStart:IMAP4ExtendedListCommand
            Using client As New ImapClient("imap.gmail.com", 993, "username", "password")
                Dim folderInfoCol As ImapFolderInfoCollection = client.ListFolders("*")
                Console.WriteLine("Extended List Supported: " + client.ExtendedListSupported)
                For Each folderInfo As ImapFolderInfo In folderInfoCol
                    Select Case folderInfo.Name
                        Case "[Gmail]/All Mail"
                            Console.WriteLine("Has Children: " + folderInfo.HasChildren)
                            Exit Select
                        Case "[Gmail]/Bin"
                            Console.WriteLine("Bin has children? " + folderInfo.HasChildren)
                            Exit Select
                        Case "[Gmail]/Drafts"
                            Console.WriteLine("Drafts has children? " + folderInfo.HasChildren)
                            Exit Select
                        Case "[Gmail]/Important"
                            Console.WriteLine("Important has Children? " + folderInfo.HasChildren)
                            Exit Select
                        Case "[Gmail]/Sent Mail"
                            Console.WriteLine("Sent Mail has Children? " + folderInfo.HasChildren)
                            Exit Select
                        Case "[Gmail]/Spam"
                            Console.WriteLine("Spam has Children? " + folderInfo.HasChildren)
                            Exit Select
                        Case "[Gmail]/Starred"
                            Console.WriteLine("Starred has Children? " + folderInfo.HasChildren)
                            Exit Select
                    End Select
                Next
            End Using
            'ExEnd:IMAP4ExtendedListCommand
        End Sub
    End Class
End Namespace