Imports System.IO
Imports System.Collections.Generic
Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class ParseSearchableFolders
        Public Shared Sub Run()
            WalkAllFolders()
        End Sub

        ' Open an OST and get the name and
        ' parent name for all folders in the file
        Private Shared Sub WalkAllFolders()
            'ExStart:ParseSearchableFolders
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim path As String = dataDir & Convert.ToString("PersonalStorage.pst")
            Dim buffer As Byte() = File.ReadAllBytes(path)
            Using s As Stream = File.OpenRead(path)
                Dim pst As PersonalStorage = PersonalStorage.FromStream(s)

                ' One sample of the sub-folder EntryId which could not be retrieved in Finder sub-folder
                Dim target As FolderInfo = pst.GetFolderById("AAAAAB9of1CGOidPhTb686WQY68igAAA")
                Dim folderData As IList(Of String) = New List(Of String)()
                WalkFolders(pst.RootFolder, "N/A", folderData)
                Dim foo = "foo"
            End Using
            'ExEnd:ParseSearchableFolders
        End Sub

        ' Walk the folder structure recursively
        Private Shared Sub WalkFolders(folder As FolderInfo, parentFolderName As String, folderData As IList(Of String))
            'ExStart:ParseSearchableFolders-WalkFolders
            Dim displayName As String = If((String.IsNullOrEmpty(folder.DisplayName)), "ROOT", folder.DisplayName)
            Dim folderNames As String = String.Format("DisplayName = {0} Parent.DisplayName = {1}", displayName, parentFolderName)
            folderData.Add(folderNames)
            If displayName = "Finder" Then
                Console.WriteLine("Test this case")
            End If

            If Not folder.HasSubFolders Then
                Return
            End If
            Dim coll As FolderInfoCollection = folder.GetSubFolders(FolderKind.Search Or FolderKind.Normal)
            For Each subfolder As FolderInfo In coll
                WalkFolders(subfolder, displayName, folderData)
            Next
            'ExEnd:ParseSearchableFolders-WalkFolders
        End Sub
    End Class
End Namespace