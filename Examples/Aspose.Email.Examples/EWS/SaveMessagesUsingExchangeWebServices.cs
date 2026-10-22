// Demonstrates saving messages straight to .eml files.
//
// SaveMessage writes the message's MIME content to a file without parsing it into a
// MailMessage first - the quickest way to archive mail. The files open in any mail
// program and load with MailMessage.Load.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class SaveMessagesUsingExchangeWebServices
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("EwsEml");
            var index = 0;

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                foreach (var info in client.ListMessages(client.MailboxInfo.InboxUri, 5))
                {
                    client.SaveMessage(info.UniqueUri, outputDir/$"message-{++index}.eml");
                    Console.WriteLine($"Saved message-{index}.eml  {info.Subject}");
                }
            }

            Console.WriteLine($"\n{index} message(s) saved to {outputDir}");
        }
    }
}
