using System;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
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
    class AddMembersToPrivateDistributionList
    {
        public static void Run()
        {
            // ExStart:AddMembersToPrivateDistributionList
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");
            ExchangeDistributionList[] distributionLists = client.ListDistributionLists();
            MailAddressCollection newMembers = new MailAddressCollection();
            newMembers.Add("address4@host.com");
            newMembers.Add("address5@host.com");
            client.AddToDistributionList(distributionLists[0], newMembers);
            // ExEnd:AddMembersToPrivateDistributionList
        }
    }
}