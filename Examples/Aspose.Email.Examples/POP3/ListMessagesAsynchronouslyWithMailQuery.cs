using Aspose.Email.Clients;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.POP3
{
    class ListMessagesAsynchronouslyWithMailQuery
    {
        public static void Run()
        {
            // Pop3Client client = new Pop3Client();
            // client.Host = "pop.gmail.com";
            // client.Port = 995;
            // client.SecurityOptions = SecurityOptions.SSLImplicit;
            // client.Username = "username";
            // client.Password = "password";
            //
            // try
            // {
            //     MailQueryBuilder builder = new MailQueryBuilder();
            //     builder.Subject.Contains("Subject");
            //     MailQuery query = builder.GetQuery();
            //     IAsyncResult asyncResult = client.BeginListMessages(query);
            //     Pop3MessageInfoCollection messages = client.EndListMessages(asyncResult);
            // }
            // catch (Exception ex)
            // {
            //     Console.WriteLine(ex.Message);
            // }
        }
    }
}
