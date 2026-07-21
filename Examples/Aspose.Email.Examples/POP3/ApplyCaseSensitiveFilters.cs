using Aspose.Email.Clients.Pop3;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.POP3
{
    class ApplyCaseSensitiveFilters
    {
        public static void Run()
        {

            // Connect and log in to POP3
            const string host = "host";
            const int port = 110;
            const string username = "user@host.com";
            const string password = "password";
            Pop3Client client = new Pop3Client(host, port, username, password);

            try
            {
                // IgnoreCase is True
                MailQueryBuilder builder1 = new MailQueryBuilder();
                builder1.From.Contains("tesT", true);
                MailQuery query1 = builder1.GetQuery();
                Pop3MessageInfoCollection messageInfoCol1 = client.ListMessages(query1);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
