Imports System.IO
Imports Aspose.Email.Outlook.Pst

' This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
'   API reference when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq 
'   for more information. If you do not wish to use NuGet, you can manually download 
'   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'   install it and then add its reference to this project. For any issues, questions or suggestions 
'   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
'

Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreateNewPSTFileAndAddingSubfolders
        Public Shared Sub Run()
            ' ExStart:CreateNewPSTFileAndAddingSubfolders
            ' The path to the file directory.
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()

            Dim checkfile As String = dataDir + "CreateNewPSTFileAndAddingSubfolders_out.pst"

            If File.Exists(checkfile) Then
                File.Delete(checkfile)

            Else
            End If

            ' Load the Outlook file
            Dim path As String = dataDir & Convert.ToString("CreateNewPSTFileAndAddingSubfolders_out.pst")

            ' Create new PST
            Dim personalStorage__1 As PersonalStorage = PersonalStorage.Create(path, FileFormatVersion.Unicode)

            ' Add new folder "Test"
            personalStorage__1.RootFolder.AddSubFolder("Inbox")
            ' ExEnd:CreateNewPSTFileAndAddingSubfolders           
        End Sub
    End Class
End Namespace
