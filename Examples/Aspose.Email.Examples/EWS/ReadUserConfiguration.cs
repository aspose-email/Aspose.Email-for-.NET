using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    class ReadUserConfiguration
    {
        public static void Run()
        {

            IEWSClient client = GetExchangeEWSClient();
            Console.WriteLine("Connected to Exchange 2010");

            // Get the User Configuration for Inbox folder
            UserConfigurationName userConfigName = new UserConfigurationName("inbox.config", client.MailboxInfo.InboxUri);
            UserConfiguration userConfig = client.GetUserConfiguration(userConfigName);

            Console.WriteLine("Configuration Id: " + userConfig.Id);
            Console.WriteLine("Configuration Name: " + userConfig.UserConfigurationName.Name);
            Console.WriteLine("Key value pairs:");
            foreach (string key in userConfig.Dictionary.Keys)
            {
                Console.WriteLine(key + ": " + userConfig.Dictionary[key].ToString());
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