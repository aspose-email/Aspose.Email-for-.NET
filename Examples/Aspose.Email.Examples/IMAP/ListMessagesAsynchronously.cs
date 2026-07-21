using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.IMAP
{
    class ListMessagesAsynchronously
    {
        // public static void Run()
        // {
        //     // Create an instance of the ImapClient class
        //     ImapClient client = new ImapClient();
        //
        //     // Specify host, username, password, port and SecurityOptions for your client
        //     client.Host = "imap.gmail.com";
        //     client.Username = "your.username@gmail.com";
        //     client.Password = "your.password";
        //     client.Port = 993;
        //     client.SecurityOptions = SecurityOptions.Auto;
        //     try
        //     {
        //         ImapQueryBuilder builder = new ImapQueryBuilder();
        //         builder.Subject.Contains("Subject");
        //         MailQuery query = builder.GetQuery();
        //         IAsyncResult asyncResult = client.BeginListMessages(query);
        //         ImapMessageInfoCollection messages = client.EndListMessages(asyncResult);
        //
        //         Console.WriteLine("New Message Added Successfully");
        //         client.Dispose();
        //     }
        //     catch (Exception ex)
        //     {
        //         Console.Write(Environment.NewLine + ex);
        //     }
        // }
    }
}
