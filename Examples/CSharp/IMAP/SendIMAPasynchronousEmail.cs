using System;
using System.Collections.Generic;
using System.Threading;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class SendIMAPasynchronousEmail
    {
        public static void Run()
        {
            try
            {

            // ExStart:SendIMAPasynchronousEmail
            // Create an imapclient with host, user and password
            ImapClient client = new ImapClient();
            client.Host = "domain.com";
            client.Username = "username";
            client.Password = "password";
            client.SelectFolder("InBox");

            ImapMessageInfoCollection messages = client.ListMessages();
            IAsyncResult res1 = client.BeginFetchMessage(messages[0].UniqueId);
            IAsyncResult res2 = client.BeginFetchMessage(messages[1].UniqueId);
            MailMessage msg1 = client.EndFetchMessage(res1);
            MailMessage msg2 = client.EndFetchMessage(res2);

            // ExEnd:SendIMAPasynchronousEmail

            // ExStart:SendIMAPasynchronousEmailThreadPool
            List<MailMessage> List = new List<MailMessage>();
            ThreadPool.QueueUserWorkItem(delegate(object o)
            {
                client.SelectFolder("folderName");
                ImapMessageInfoCollection messageInfoCol = client.ListMessages();
                foreach (ImapMessageInfo messageInfo in messageInfoCol)
                {
                    List.Add(client.FetchMessage(messageInfo.UniqueId));
                }
            });
            // ExEnd:SendIMAPasynchronousEmailThreadPool

            // ExStart:SendIMAPasynchronousEmailThreadPool1
            List<MailMessage> List1 = new List<MailMessage>();
            ThreadPool.QueueUserWorkItem(delegate(object o)
            {
                using (IDisposable connection = client.CreateConnection())
                {
                    client.SelectFolder("FolderName");
                    ImapMessageInfoCollection messageInfoCol =
               client.ListMessages();
                    foreach (ImapMessageInfo messageInfo in messageInfoCol)
                        List1.Add(client.FetchMessage(messageInfo.UniqueId));
                }
            });
            // ExEnd:SendIMAPasynchronousEmailThreadPool1
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
                throw;
            }
        }
    }
}