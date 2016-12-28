using Aspose.Email.Outlook;
using Aspose.Email.Outlook.Pst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Email
{
    class Program
    {
        static void Main(string[] args)
        {
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
            task.Save("MapiTask.msg", TaskSaveFormat.Msg);
        }
    }
}
