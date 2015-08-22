using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace VSTO_Outlook
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            CreateRecurringTask();
        }

        private void CreateRecurringTask()
        {
            Outlook.TaskItem task = Application.CreateItem(
                Outlook.OlItemType.olTaskItem) as Outlook.TaskItem;
            task.Subject = "Tax Preparation";
            task.StartDate = DateTime.Parse("4/1/2007 8:00 AM");
            task.DueDate = DateTime.Parse("4/15/2007 8:00 AM");
            Outlook.RecurrencePattern pattern =
                task.GetRecurrencePattern();
            pattern.RecurrenceType = Outlook.OlRecurrenceType.olRecursYearly;
            pattern.PatternStartDate = DateTime.Parse("4/1/2007");
            pattern.NoEndDate = true;
            task.ReminderSet = true;
            task.ReminderTime = DateTime.Parse("4/1/2007 8:00 AM");
            task.Save();
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
