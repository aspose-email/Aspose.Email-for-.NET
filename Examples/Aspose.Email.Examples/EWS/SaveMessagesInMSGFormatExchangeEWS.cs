// Demonstrates saving messages from Exchange as Outlook .msg files.
//
// FetchMessage downloads the message as a MailMessage and Save converts it to the
// format chosen with SaveOptions - here Unicode MSG, which Outlook opens directly.
// SaveMessagesUsingExchangeWebServices keeps the original MIME data instead.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class SaveMessagesInMSGFormatExchangeEWS
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("EwsMsg");
            var index = 0;

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                foreach (var info in client.ListMessages(client.MailboxInfo.InboxUri, 5))
                {
                    var message = client.FetchMessage(info.UniqueUri);
                    message.Save(outputDir/$"message-{++index}.msg", SaveOptions.DefaultMsgUnicode);
                    Console.WriteLine($"Saved message-{index}.msg  {message.Subject}");
                }
            }

            Console.WriteLine($"\n{index} message(s) saved to {outputDir}");
        }
    }
}
