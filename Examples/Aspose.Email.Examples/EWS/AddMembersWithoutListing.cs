using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    class AddMembersWithoutListing
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {

                // Create a local distribution list with a known Id and ChangeKey.
                ExchangeDistributionList distributionList =  new ExchangeDistributionList()
                {
                    Id = "AAMkAGJhZjYzY2I5LTdjYWMtNGFmMC05ODI1LTA5MTAzYTgwZTc4OQBGAAAAAABdN1MC60QcSpWwPYUTPhL2BwATlR+p0q0wT6WD0+d4WJhWAAAAAAEOAAATlR+p0q0wT6WD0+d4WJhWAACImuRlAAA=",
                    ChangeKey = "EgAAABYAAAATlR+p0q0wT6WD0+d4WJhWAACKy8ZX"
                };

                // Init a new members
                MailAddressCollection newMembers = new MailAddressCollection()
                {
                    "address6@host.com"
                };

                // Udpdate a distribution list on the server with new member.
                client.AddToDistributionList(distributionList, newMembers);
            }
        }
    }
}