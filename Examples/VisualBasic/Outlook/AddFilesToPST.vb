Imports System.IO
Imports Aspose.Email.Outlook.Pst

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class AddFilesToPST
        Public Shared Sub Run()
            ' The path to the File directory.
            ' ExStart:AddFilesToPST
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim checkfile As String = dataDir + "Ps1_out.pst"
            If File.Exists(checkfile) Then
                File.Delete(checkfile)
            Else
            End If


            Using personalStorage1 As PersonalStorage = PersonalStorage.Create(dataDir & Convert.ToString("Ps1_out.pst"), FileFormatVersion.Unicode)
                Dim folder As FolderInfo = personalStorage1.RootFolder.AddSubFolder("Files")

                ' Add Document.doc file with the "IPM.Document" message class by default.
                folder.AddFile(dataDir & Convert.ToString("attachment_1.doc"), Nothing)
            End Using
            ' ExEnd:AddFilesToPST
        End Sub
    End Class
End Namespace