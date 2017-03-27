using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.IMAP
{
    class ListMessagesAsynchronously
    {
        public static void Run()
        {
            // Create an instance of the ImapClient class
            ImapClient client = new ImapClient();

            // Specify host, username, password, port and SecurityOptions for your client
            client.Host = "imap.gmail.com";
            client.Username = "your.username@gmail.com";
            client.Password = "your.password";
            client.Port = 993;
            client.SecurityOptions = SecurityOptions.Auto;
            try
            {
                // ExStart:ListMessagesAsynchronouslyWithMailQuery
                ImapQueryBuilder builder = new ImapQueryBuilder();
                builder.Subject.Contains("Subject");
                MailQuery query = builder.GetQuery();
                IAsyncResult asyncResult = client.BeginListMessages(query);
                ImapMessageInfoCollection messages = client.EndListMessages(asyncResult);
                // ExEnd:ListMessagesAsynchronouslyWithMailQuery

                Console.WriteLine("New Message Added Successfully");
                client.Dispose();
            }
            catch (Exception ex)
            {
                Console.Write(Environment.NewLine + ex);
            }
        }
    }
}
