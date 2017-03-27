using Aspose.Email.Clients;
using Aspose.Email.Clients.Pop3;
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

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class ListMessagesAsynchronouslyWithMailQuery
    {
        public static void Run()
        {
            Pop3Client client = new Pop3Client();
            client.Host = "pop.gmail.com";
            client.Port = 995;
            client.SecurityOptions = SecurityOptions.SSLImplicit;
            client.Username = "username";
            client.Password = "password";
            
            try
            {
                // ExStart:ListMessagesAsynchronouslyWithMailQuery
                MailQueryBuilder builder = new MailQueryBuilder();
                builder.Subject.Contains("Subject");
                MailQuery query = builder.GetQuery();
                IAsyncResult asyncResult = client.BeginListMessages(query);
                Pop3MessageInfoCollection messages = client.EndListMessages(asyncResult);
                // ExEnd:ListMessagesAsynchronouslyWithMailQuery
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
