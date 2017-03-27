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
    class CreateREAndFWMessages
    {
        public static void Run()
        {
            // ExStart:CreateREAndFWMessages
            string dataDir = RunExamples.GetDataDir_Exchange();

            const string mailboxUri = "https://exchange.domain.com/ews/Exchange.asmx";
            const string domain = @"";
            const string username = @"username";
            const string password = @"password";
            NetworkCredential credential = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credential);

            try
            {
                MailMessage message = new MailMessage("user@domain.com", "user@domain.com", "TestMailRefw - " + Guid.NewGuid().ToString(),
               "TestMailRefw Implement ability to create RE and FW messages from source MSG file");

                client.Send(message);

                ExchangeMessageInfoCollection messageInfoCol = client.ListMessages(client.MailboxInfo.InboxUri);
                if (messageInfoCol.Count == 1)
                    Console.WriteLine("1 message in Inbox");
                else
                    Console.WriteLine("Error! No message in Inbox");

                MailMessage message1 = new MailMessage("user@domain.com", "user@domain.com", "TestMailRefw - " + Guid.NewGuid().ToString(),
                "TestMailRefw Implement ability to create RE and FW messages from source MSG file");

                client.Send(message1);
                messageInfoCol = client.ListMessages(client.MailboxInfo.InboxUri);

                if (messageInfoCol.Count == 2)
                    Console.WriteLine("2 messages in Inbox");
                else
                    Console.WriteLine("Error! No new message in Inbox");
                
                MailMessage message2 = new MailMessage("user@domain.com", "user@domain.com", "TestMailRefw - " + Guid.NewGuid().ToString(),
                "TestMailRefw Implement ability to create RE and FW messages from source MSG file");
                message2.Attachments.Add(Attachment.CreateAttachmentFromString("Test attachment 1", "Attachment Name 1"));
                message2.Attachments.Add(Attachment.CreateAttachmentFromString("Test attachment 2", "Attachment Name 2"));

                // Reply, Replay All and forward Message
                client.Reply(message2, messageInfoCol[0]);
                client.ReplyAll(message2, messageInfoCol[0]);
                client.Forward(message2, messageInfoCol[0]);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in program"+ex.Message);
            }
            // ExEnd:CreateREAndFWMessages
        }
    }
}