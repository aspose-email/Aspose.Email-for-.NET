using System;

namespace Aspose.Email.Examples.EWS
{
    class ConnectingToExchangeServerUsingEWS
    {
        public static void Run()
        {              
            try
            {
                var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission);
                Console.WriteLine(client.ListMessages(client.MailboxInfo.InboxUri).Count);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
