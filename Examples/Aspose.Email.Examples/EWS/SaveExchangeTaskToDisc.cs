// Demonstrates saving an Exchange task to a file without a server.
//
// ExchangeTask is the task model of the EWS client. Save writes it to disk - in
// Outlook's .msg format by default - so tasks read from Exchange can be archived, or
// tasks prepared offline can be opened in Outlook. This example runs offline.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class SaveExchangeTaskToDisc
    {
        public static void Run()
        {
            var outputPath = Data.Out/"SaveExchangeTaskToDisc_out.msg";

            var task = new ExchangeTask
            {
                Subject = "Prepare the quarterly report",
                Status = ExchangeTaskStatus.InProgress,
                StartDate = DateTime.Today,
                DueDate = DateTime.Today.AddDays(3),
                Body = "Collect the numbers from all regions."
            };

            task.Save(outputPath);

            Console.WriteLine($"Saved '{task.Subject}' ({task.Status}, due {task.DueDate:d}) to {outputPath}");
        }
    }
}
