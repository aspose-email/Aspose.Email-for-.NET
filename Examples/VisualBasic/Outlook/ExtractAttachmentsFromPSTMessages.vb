Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class ExtractAttachmentsFromPSTMessages
        Public Shared Sub Run()
            ' The path to the File directory.
            ' ExStart:ExtractAttachmentsFromPSTMessages
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Using personalstorage__1 As PersonalStorage = PersonalStorage.FromFile(dataDir & Convert.ToString("Outlook_1.pst"))
                Dim folder As FolderInfo = personalstorage__1.RootFolder.GetSubFolder("Inbox")

                For Each messageInfo In folder.EnumerateMessagesEntryId()
                    Dim attachments As MapiAttachmentCollection = personalstorage__1.ExtractAttachments(messageInfo)

                    If attachments.Count <> 0 Then
                        For Each attachment In attachments
                            If Not String.IsNullOrEmpty(attachment.LongFileName) Then
                                If attachment.LongFileName.Contains(".msg") Then
                                    Continue For
                                Else
                                    attachment.Save((dataDir & Convert.ToString("\Attachments\")) + attachment.LongFileName.ToString())
                                End If
                            End If
                        Next
                    End If
                Next
            End Using
            ' ExEnd:ExtractAttachmentsFromPSTMessages
        End Sub
    End Class
End Namespace
