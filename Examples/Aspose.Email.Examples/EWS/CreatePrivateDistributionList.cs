using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class CreatePrivateDistributionList
    {       
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {

                var distributionList = new ExchangeDistributionList
                {
                    DisplayName = "test private list"
                };

                var members = new MailAddressCollection
                {
                    "address1@host.com",
                    "address2@host.com",
                    "address3@host.com"
                };

                client.CreateDistributionList(distributionList, members);
            }
        }       
    }
}