Imports System.IO
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
    Class AddMapiTaskToPST
        Public Shared Sub Run()
            ' The path to the documents directory.
            ' ExStart:AddMapiTaskToPST
            Dim dataDir As String = RunExamples.GetDataDir_Outlook()
            Dim task As New MapiTask("To Do", "Just click and type to add new task", DateTime.Now, DateTime.Now.AddDays(3))
            task.PercentComplete = 20
            task.EstimatedEffort = 2000
            task.ActualEffort = 20
            task.History = MapiTaskHistory.Assigned
            task.LastUpdate = DateTime.Now
            task.Users.Owner = "Darius"
            task.Users.LastAssigner = "Harkness"
            task.Users.LastDelegate = "Harkness"
            task.Users.Ownership = MapiTaskOwnership.AssignersCopy

            Dim checkfile As String = dataDir + "AddMapiTaskToPST_out.pst"

            If File.Exists(checkfile) Then
                File.Delete(checkfile)

            Else
            End If
            Using personalStorage__1 As PersonalStorage = PersonalStorage.Create(dataDir & Convert.ToString("AddMapiTaskToPST_out.pst"), FileFormatVersion.Unicode)
                Dim taskFolder As FolderInfo = personalStorage__1.CreatePredefinedFolder("Tasks", StandardIpmFolder.Tasks)
                taskFolder.AddMapiMessageItem(task)
            End Using
            ' ExEnd:AddMapiTaskToPST
        End Sub
    End Class
End Namespace
