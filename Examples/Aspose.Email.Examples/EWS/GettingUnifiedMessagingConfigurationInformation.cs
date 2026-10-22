// Demonstrates reading the Unified Messaging (voice mail) settings of a mailbox.
//
// GetUMConfiguration tells whether Unified Messaging is enabled and whether "Play on
// Phone" can read messages out over a phone call. Unified Messaging was removed in
// Exchange Server 2019 and Exchange Online, where it reports as disabled; the call is
// meaningful for Exchange 2010 to 2016.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class GettingUnifiedMessagingConfigurationInformation
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var configuration = client.GetUMConfiguration();

                Console.WriteLine($"Unified Messaging enabled: {configuration.UMEnabled}");
                Console.WriteLine($"Play on Phone enabled:     {configuration.PlayOnPhoneEnabled}");
                Console.WriteLine($"Play on Phone dial string: {configuration.PlayOnPhoneDialString}");
            }
        }
    }
}
