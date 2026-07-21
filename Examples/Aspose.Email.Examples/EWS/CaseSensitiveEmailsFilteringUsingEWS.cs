using Aspose.Email.Clients.Exchange;
using System;

namespace Aspose.Email.Examples.EWS
{
    class CaseSensitiveEmailsFilteringUsingEWS
    {
        public static void Run()
        {
            try
            {
                using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
                {
                    // Query building by means of ExchangeQueryBuilder class
                    var builder = new ExchangeQueryBuilder();
                    builder.Subject.Contains("Newsletter", true);
                    builder.InternalDate.On(DateTime.Now);
                    var query = builder.GetQuery();

                    // Get list of messages
                    var messages = client.ListMessages(client.MailboxInfo.InboxUri, query, false);
                    Console.WriteLine($"EWS: {messages.Count} message(s) found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}

