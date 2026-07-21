using System;
using System.Collections.Generic;
using System.Threading;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.IMAP
{
    class SendIMAPasynchronousEmail
    {
        public static void Run()
        {
        //     try
        //     {
        //
        //     // Create an imapclient with host, user and password
        //     ImapClient client = new ImapClient();
        //     client.Host = "domain.com";
        //     client.Username = "username";
        //     client.Password = "password";
        //     client.SelectFolder("InBox");
        //
        //     ImapMessageInfoCollection messages = client.ListMessages();
        //     IAsyncResult res1 = client.BeginFetchMessage(messages[0].UniqueId);
        //     IAsyncResult res2 = client.BeginFetchMessage(messages[1].UniqueId);
        //     MailMessage msg1 = client.EndFetchMessage(res1);
        //     MailMessage msg2 = client.EndFetchMessage(res2);
        //
        //
        //     List<MailMessage> List = new List<MailMessage>();
        //     ThreadPool.QueueUserWorkItem(delegate(object o)
        //     {
        //         client.SelectFolder("folderName");
        //         ImapMessageInfoCollection messageInfoCol = client.ListMessages();
        //         foreach (ImapMessageInfo messageInfo in messageInfoCol)
        //         {
        //             List.Add(client.FetchMessage(messageInfo.UniqueId));
        //         }
        //     });
        //
        //     List<MailMessage> List1 = new List<MailMessage>();
        //     ThreadPool.QueueUserWorkItem(delegate(object o)
        //     {
        //         using (IDisposable connection = client.CreateConnection())
        //         {
        //             client.SelectFolder("FolderName");
        //             ImapMessageInfoCollection messageInfoCol =
        //        client.ListMessages();
        //             foreach (ImapMessageInfo messageInfo in messageInfoCol)
        //                 List1.Add(client.FetchMessage(messageInfo.UniqueId));
        //         }
        //     });
        //     }
        //     catch (Exception ex)
        //     {
        //         Console.Write(ex.Message);
        //         throw;
        //     }
         }
    }
}