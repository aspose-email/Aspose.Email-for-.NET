// Demonstrates finding out which Exchange version the client is talking to.
//
// Some EWS operations exist only on newer servers - FindPeople needs Exchange 2013, for
// example. GetExchangeType and ServerVersion return the version as an ExchangeVersion
// value to branch on; GetVersionInfo returns the server's own version string, which is
// handy for logs and support requests.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class GetExchangeServerVersion
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var version = client.GetExchangeType();

                Console.WriteLine($"Exchange version: {version}");
                Console.WriteLine($"ServerVersion:    {client.ServerVersion}");
                Console.WriteLine($"Version info:     {client.GetVersionInfo()}");

                Console.WriteLine(version >= ExchangeVersion.Exchange2013
                    ? "\nFindPeople and the other Exchange 2013 operations are available."
                    : "\nThis server predates Exchange 2013; FindPeople is not available.");
            }
        }
    }
}
