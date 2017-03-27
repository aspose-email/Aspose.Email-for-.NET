using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class UpdateTaskOnExchange
    {
        public static void Run()
        {

            try
            {

                // ExStart:UpdateTaskOnExchange
                // Create and initialize credentials
                var credentials = new NetworkCredential("username", "12345");

                // Create instance of ExchangeClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // Get all tasks info collection from exchange
                ExchangeMessageInfoCollection tasks = client.ListMessages(client.MailboxInfo.TasksUri);

                // Parse all the tasks info in the list
                foreach (ExchangeMessageInfo info in tasks)
                {
                    // Fetch task from exchange using current task info
                    ExchangeTask task = client.FetchTask(info.UniqueUri);

                    // Update the task status to NotStarted
                    task.Status = ExchangeTaskStatus.NotStarted;

                    // Set the task due date
                    task.DueDate = new DateTime(2013, 2, 26);

                    // Set task priority
                    task.Priority = MailPriority.Low;

                    // Update task on exchange
                    client.UpdateTask(task);
                }
                // ExEnd:UpdateTaskOnExchange
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
    }
}