using System;

namespace Aspose.Email.Examples.EWS
{
    class CopyingMessageToAnotherFolder
    {       
        public static void Run()
        {
            try
            {
                using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
                {
                    var message = new MailMessage("from@domain.com", "to@domain.com", "Notes", "This message contains release notes.");
                    // Append message to the Inbox folder
                    var messageUri = client.AppendMessage(message);
                    // Copy message to Deleted Items
                    var newMessageUri = client.CopyItem(messageUri, client.MailboxInfo.DeletedItemsUri);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }       
    }
}