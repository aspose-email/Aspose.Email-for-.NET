using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    class UpdateUserConfiguration
    {
        public static void Run()
        {
            try
            {

                IEWSClient client = GetExchangeEWSClient();
                Console.WriteLine("Connected to Exchange 2010");

                // Create the User Configuration for Inbox folder
                UserConfigurationName userConfigName = new UserConfigurationName("inbox.config", client.MailboxInfo.InboxUri);
                UserConfiguration userConfig = client.GetUserConfiguration(userConfigName);
                userConfig.Id = null;

                // Update User Configuration
                userConfig.Dictionary["key1"] = "new-value1";
                client.UpdateUserConfiguration(userConfig);
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }

        private static IEWSClient GetExchangeEWSClient()
        {
            const string mailboxUri = "https://outlook.office365.com/ews/exchange.asmx";
            const string domain = @"";
            const string username = @"username@ASE305.onmicrosoft.com";
            const string password = @"password";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);
            return client;
        }
    }
}