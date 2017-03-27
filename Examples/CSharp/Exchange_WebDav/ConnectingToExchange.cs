using Aspose.Email.Clients.Exchange.Dav;
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

namespace Aspose.Email.Examples.CSharp.Email.Exchange_WebDav
{
    class ConnectingToExchange
    {
        public static void Run()
        {
            try
            {
                ExchangeClient client = GetExchangeClient();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        // ExStart:ConnectingToExchange
        static ExchangeClient GetExchangeClient()
        {
            // Create instance of ExchangeClient class by giving credentials
            ExchangeClient client = new ExchangeClient("http://MachineName/exchange/Username", "Username", "password", "domain");
            return client;
        }
        // ExEnd:ConnectingToExchange
    }
}
