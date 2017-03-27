using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net;
using System.Threading;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class CreatePrivateDistributionList
    {       
        public static void Run()
        {
            // ExStart:CreatePrivateDistributionList
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");
            ExchangeDistributionList distributionList = new ExchangeDistributionList();
            distributionList.DisplayName = "test private list";
            MailAddressCollection members = new MailAddressCollection();
            members.Add("address1@host.com");
            members.Add("address2@host.com");
            members.Add("address3@host.com");
            client.CreateDistributionList(distributionList, members);
            // ExEnd:CreatePrivateDistributionList
        }       
    }
}