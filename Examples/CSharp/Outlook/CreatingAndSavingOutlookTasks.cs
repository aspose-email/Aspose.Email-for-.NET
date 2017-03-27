using System;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Mapi;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class CreatingAndSavingOutlookTasks
    {
        public static void Run()
        {
            // ExStart:CreatingAndSavingOutlookTasks
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            MapiTask task = new MapiTask("To Do", "Just click and type to add new task", DateTime.Now, DateTime.Now.AddDays(3));
            task.PercentComplete = 20;
            task.EstimatedEffort = 2000;
            task.ActualEffort = 20;
            task.History = MapiTaskHistory.Assigned;
            task.LastUpdate = DateTime.Now;
            task.Users.Owner = "Darius";
            task.Users.LastAssigner = "Harkness";
            task.Users.LastDelegate = "Harkness";
            task.Users.Ownership = MapiTaskOwnership.AssignersCopy;
            task.Companies = new string[] { "company1", "company2", "company3" };
            task.Categories = new string[] { "category1", "category2", "category3" };
            task.Mileage = "Some test mileage";
            task.Billing = "Test billing information";
            task.Users.Delegator = "Test Delegator";
            task.Sensitivity = MapiSensitivity.Personal;
            task.Status = MapiTaskStatus.Complete;
            task.EstimatedEffort = 5;
            task.Save(dataDir + "MapiTask.msg", TaskSaveFormat.Msg);
            // ExEnd:CreatingAndSavingOutlookTasks
        }
    }
}