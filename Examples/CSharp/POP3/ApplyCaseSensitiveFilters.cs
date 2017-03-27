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
                // ExStart:ApplyCaseSensitiveFilters
                // IgnoreCase is True
                MailQueryBuilder builder1 = new MailQueryBuilder();
                builder1.From.Contains("tesT", true);
                MailQuery query1 = builder1.GetQuery();
                Pop3MessageInfoCollection messageInfoCol1 = client.ListMessages(query1);
                // ExEnd:ApplyCaseSensitiveFilters
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
