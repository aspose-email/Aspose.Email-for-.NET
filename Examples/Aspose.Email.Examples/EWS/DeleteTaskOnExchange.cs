using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class DeleteTaskOnExchange
    {
        public static void Run()
        {
            // Create instance of ExchangeClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            // Get all tasks info collection from exchange
            ExchangeMessageInfoCollection tasks = client.ListMessages(client.MailboxInfo.TasksUri);

            // Parse all the tasks info in the list
            foreach (ExchangeMessageInfo info in tasks)
            {
                // Fetch task from exchange using current task info
                ExchangeTask task = client.FetchTask(info.UniqueUri);

                // Check if the current task fulfills the search criteria
                if (task.Subject.Equals("test"))
                {
                    //Delete task from exchange
                    client.DeleteItem(task.UniqueUri, DeletionOptions.DeletePermanently);
                }
            }
        }
    }
}