using Aspose.Email.Clients.Exchange;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Aspose.Email.Examples.EWS
{
    public class FilterMessagesOnCriteriaWithAqsUsingEWS
    {
        public static void Run()
        {
            // using the builder
            {
                using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
                {
                    List<string> messagesIDs = new List<string>();
                    var foldersList = client.ListSubFolders(client.GetMailboxInfo().RootUri);
                    var folderInfo =
                        foldersList.FirstOrDefault(folder => folder.DisplayName.Equals("test", StringComparison.Ordinal));
                    if (folderInfo == null)
                    {
                        folderInfo = client.GetFolderInfo("test");
                    }

                    ExchangeFolderInfo userFolder = folderInfo;
                    ExchangeMessageInfoCollection list = client.ListMessages(userFolder.Uri);
                    foreach (var email in list)
                    {
                        messagesIDs.Add(email.UniqueUri);
                    }
                }

                //var advancedBuilder = new ExchangeAdvancedSyntaxQueryBuilder();
                //advancedBuilder.From.Equals("Microsoft");
                //advancedBuilder.Subject.Contains("credit card");

                //var msgCollection = ewsClient.ListMessages(ewsClient.MailboxInfo.InboxUri, advancedBuilder.GetQuery());

                //Console.WriteLine(@"Filter messages by using the query builder: ");
                //Console.WriteLine(@"=========================================");

                //foreach (var msgInfo in msgCollection)
                //{
                //    Console.WriteLine(msgInfo.Subject);
                //}
            }

            //Console.WriteLine();

            //// using the AQS
            //{
            //    using var ewsClient = ClientBuilder.Ews(AuthType.ModernWithAppPermission);

            //    ExchangeAdvancedSyntaxMailQuery query = new ExchangeAdvancedSyntaxMailQuery("subject:Your digest email");

            //    var msgCollection = ewsClient.ListMessages(ewsClient.MailboxInfo.InboxUri, query);

            //    Console.WriteLine(@"Filter messages by using the AQS syntax: ");
            //    Console.WriteLine(@"=========================================");

            //    foreach (var msgInfo in msgCollection)
            //    {
            //        Console.WriteLine(msgInfo.Subject);
            //    }
            //}
        }
    }
}