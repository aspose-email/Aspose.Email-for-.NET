Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class GetMessageInformation
        Public Shared Sub Run()
            ' ExStart:GetMessageInformation
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            ' Load the Outlook file
            Dim path As String = dataDir & Convert.ToString("PersonalStorage.pst")

            Try
                ' Load the Outlook PST file
                Dim personalStorage__1 As PersonalStorage = PersonalStorage.FromFile(path)

                ' Get the Display Format of the PST file
                Console.WriteLine("Display Format: " + personalStorage__1.Format.ToString())

                ' Get the folders and messages information
                Dim folderInfo As FolderInfo = personalStorage__1.RootFolder

                ' Call the recursive method to display the folder contents
                DisplayFolderContents(folderInfo, personalStorage__1)
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
            ' ExEnd:GetMessageInformation
        End Sub

        ' <summary>
        ' This is a recursive method to display contents of a folder
        ' </summary>
        ' <param name="folderInfo"></param>
        ' <param name="pst"></param>
        Private Shared Sub DisplayFolderContents(folderInfo As FolderInfo, pst As PersonalStorage)
            ' ExStart:GetMessageInformationDisplayFolderContents
            ' Display the folder name
            Console.WriteLine("Folder: " + folderInfo.DisplayName.ToString())
            ' Display information about messages inside this folder
            Dim messageInfoCollection As MessageInfoCollection = folderInfo.GetContents()
            For Each messageInfo As MessageInfo In messageInfoCollection
                Console.WriteLine("Subject: " + messageInfo.Subject.ToString())
                Console.WriteLine("Sender: " + messageInfo.SenderRepresentativeName.ToString())
                Console.WriteLine("Recipients: " + messageInfo.DisplayTo.ToString())
            Next

            ' Call this method recursively for each subfolder
            If folderInfo.HasSubFolders = True Then
                For Each subfolderInfo As FolderInfo In folderInfo.GetSubFolders()
                    DisplayFolderContents(subfolderInfo, pst)
                Next
            End If
            ' ExEnd:GetMessageInformationDisplayFolderContents
        End Sub
    End Class
End Namespace