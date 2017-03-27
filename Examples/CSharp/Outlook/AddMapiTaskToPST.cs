using System;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System.IO;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class AddMapiTaskToPST
    {
        public static void Run()
        {
            // The path to the File directory.
            // ExStart:AddMapiTaskToPST
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

            string alreadyCreated = dataDir + "AddMapiTaskToPST_out.pst";
            if (File.Exists(alreadyCreated))
            {
                File.Delete(alreadyCreated);
            }
            else
            {

            }

            using (PersonalStorage personalStorage = PersonalStorage.Create(dataDir + "AddMapiTaskToPST_out.pst", FileFormatVersion.Unicode))
            {
                FolderInfo taskFolder = personalStorage.CreatePredefinedFolder("Tasks", StandardIpmFolder.Tasks);
                taskFolder.AddMapiMessageItem(task);
            }
            // ExEnd:AddMapiTaskToPST
        }
    }
}