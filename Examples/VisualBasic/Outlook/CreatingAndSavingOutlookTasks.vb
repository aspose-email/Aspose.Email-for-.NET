Imports Aspose.Email.Outlook.Pst
Imports Aspose.Email.Outlook

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'


Namespace Aspose.Email.Examples.VisualBasic.Email.Outlook
    Class CreatingAndSavingOutlookTasks
        Public Shared Sub Run()
            ' ExStart:CreatingAndSavingOutlookTasks
            ' The path to the documents directory.
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
            task.Companies = New String() {"company1", "company2", "company3"}
            task.Categories = New String() {"category1", "category2", "category3"}
            task.Mileage = "Some test mileage"
            task.Billing = "Test billing information"
            task.Users.Delegator = "Test Delegator"
            task.Sensitivity = MapiSensitivity.Personal
            task.Status = MapiTaskStatus.Complete
            task.EstimatedEffort = 5
            task.Save(dataDir & Convert.ToString("MapiTask.msg"), TaskSaveFormat.Msg)
            ' ExEnd:CreatingAndSavingOutlookTasks
        End Sub
    End Class
End Namespace
