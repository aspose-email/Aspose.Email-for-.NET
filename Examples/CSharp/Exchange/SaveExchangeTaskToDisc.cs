using System;
using Aspose.Email.Exchange;

namespace Aspose.Email.Examples.CSharp.Email.Exchange
{
    class SaveExchangeTaskToDisc
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Exchange();
            string dstEmail = dataDir + "Message.msg";

            ExchangeTask task = new ExchangeTask();
            task.Subject = "TASK-ID - " + Guid.NewGuid();
            task.Status = ExchangeTaskStatus.InProgress;
            task.StartDate = DateTime.Now;
            task.DueDate = task.StartDate.AddDays(3);
            task.Save(dstEmail);

            Console.WriteLine(Environment.NewLine + "Task saved on disk successfully at " + dstEmail);
        }
    }
}
