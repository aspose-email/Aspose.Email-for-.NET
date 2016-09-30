Imports System.IO
Imports Aspose.Email.Outlook
Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class ExtractMessagesFromPSTFile
        Public Shared Sub Run()
            ' ExStart:ExtractMessagesFromPSTFile
            ' The path to the file directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load the Outlook file
            Dim path As String = dataDir & Convert.ToString("PersonalStorage.pst")

            Try
                ' load the Outlook PST file
                Dim pst As PersonalStorage = PersonalStorage.FromFile(path)

                ' get the Display Format of the PST file
                Console.WriteLine("Display Format: " + pst.Format.ToString())

                ' get the folders and messages information
                Dim folderInfo As FolderInfo = pst.RootFolder

                ' Call the recursive method to extract msg files from each folder
                ExtractMsgFiles(folderInfo, pst)
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
            ' ExEnd:ExtractMessagesFromPSTFile
        End Sub

        ' <summary>
        ' This is a recursive method to display contents of a folder
        ' </summary>
        ' <param name="folderInfo"></param>
        ' <param name="pst"></param>
        Private Shared Sub ExtractMsgFiles(folderInfo As FolderInfo, pst As PersonalStorage)
            ' ExStart:ExtractMessagesFromPSTFileExtractMsgFiles
            ' display the folder name
            Console.WriteLine("Folder: " + folderInfo.DisplayName)
            Console.WriteLine("==================================")
            ' loop through all the messages in this folder
            Dim messageInfoCollection As MessageInfoCollection = folderInfo.GetContents()
            For Each messageInfo As MessageInfo In messageInfoCollection
                Console.WriteLine("Saving message {0} ....", messageInfo.Subject)
                ' get the message in MapiMessage instance
                Dim message As MapiMessage = pst.ExtractMessage(messageInfo)
                ' save this message to disk in msg format
                message.Save(message.Subject.Replace(":", " ") + ".msg")
                ' save this message to stream in msg format
                Dim messageStream As New MemoryStream()
                message.Save(messageStream)
            Next

            ' Call this method recursively for each subfolder
            If folderInfo.HasSubFolders = True Then
                For Each subfolderInfo As FolderInfo In folderInfo.GetSubFolders()
                    ExtractMsgFiles(subfolderInfo, pst)
                Next
            End If
            ' ExEnd:ExtractMessagesFromPSTFileExtractMsgFiles
        End Sub
    End Class
End Namespace
