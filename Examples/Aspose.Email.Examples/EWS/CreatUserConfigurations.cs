using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    class CreatUserConfigurations
    {
        public static void Run()
        {

            IEWSClient client = GetExchangeEWSClient();
            Console.WriteLine("Connected to Exchange 2010");

            // Create the User Configuration for Inbox folder
            UserConfigurationName userConfigName = new UserConfigurationName("inbox.config", client.MailboxInfo.InboxUri);
            UserConfiguration userConfig = new UserConfiguration(userConfigName);
            userConfig.Dictionary.Add("key1", "value1");
            userConfig.Dictionary.Add("key2", "value2");
            userConfig.Dictionary.Add("key3", "value3");
            client.CreateUserConfiguration(userConfig);
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