using Aspose.Email.Clients.Exchange.WebService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class GettingUnifiedMessagingConfigurationInformation
    {
        public static void Run()
        {
            const string mailboxUri = "https://exchange.domain.com/ews/Exchange.asmx";
            const string domain = @"";
            const string username = @"username";
            const string password = @"password";
            NetworkCredential credential = new NetworkCredential(username, password, domain);

            try
            {
                // ExStart:GettingUnifiedMessagingConfigurationInformation
                IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credential);
                UnifiedMessagingConfiguration umConf = client.GetUMConfiguration();
                // ExEnd:GettingUnifiedMessagingConfigurationInformation
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
