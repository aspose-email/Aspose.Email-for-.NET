// Demonstrates finding the EWS address of a mailbox with Autodiscover.
//
// Instead of hard-coding the EWS URL, ask the Autodiscover service for it: given an
// e-mail address and credentials, GetUserSettings returns the requested settings -
// here the external EWS URL and the user's display name. This targets on-premises
// Exchange with Autodiscover; Exchange Online uses the fixed URL
// https://outlook.office365.com/EWS/Exchange.asmx. Replace the account details below.

using System;
using System.Net;
using Aspose.Email.Clients.Exchange;

namespace Aspose.Email.Examples.EWS
{
    internal static class AutoDiscoverUsingEWS
    {
        public static void Run()
        {
            const string email = "user@example.com";
            const string password = "<PASSWORD>";

            var service = new AutodiscoverService
            {
                Credentials = new NetworkCredential(email, password)
            };

            var response = service.GetUserSettings(email,
                UserSettingName.ExternalEwsUrl, UserSettingName.UserDisplayName);

            string ewsUrl;
            string displayName;
            response.TryGetSettingValue(UserSettingName.ExternalEwsUrl, out ewsUrl);
            response.TryGetSettingValue(UserSettingName.UserDisplayName, out displayName);

            Console.WriteLine($"User:     {displayName} <{response.SmtpAddress}>");
            Console.WriteLine($"EWS URL:  {ewsUrl ?? "(not returned)"}");
        }
    }
}
