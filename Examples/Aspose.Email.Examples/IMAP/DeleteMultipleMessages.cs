// Demonstrates how to delete several messages with one call.
//
// AppendMessages uploads a batch and reports the unique id of every message it stored;
// DeleteMessages then marks those ids as deleted. With UIDPLUS (RFC 4315), commitNow:
// true expunges exactly those messages at once; without it, CommitDeletes expunges
// everything marked in the folder.
//
// The messages go to a uniquely named folder, which is deleted at the end.

using System;
using System.Collections.Generic;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class DeleteMultipleMessages
    {
        public static void Run()
        {
            var folderName = "Aspose-" + Guid.NewGuid().ToString("N").Substring(0, 8);

            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                client.CreateFolder(folderName);

                try
                {
                    var messages = new List<MailMessage>();
                    for (var i = 1; i <= 5; i++)
                    {
                        messages.Add(new MailMessage("from@example.com", "to@example.com",
                            $"Message to delete {i}", "This message will be deleted."));
                    }

                    var appended = (AppendMessagesFromMessageObjectResult)client.AppendMessages(folderName, messages);
                    Console.WriteLine($"Appended {appended.Succeeded.Count} message(s), " +
                                      $"{appended.Failed.Count} failed.");

                    client.SelectFolder(folderName);
                    Console.WriteLine($"Before: {client.ListMessages().Count} message(s) in '{folderName}'");

                    if (client.UidPlusSupported)
                    {
                        client.DeleteMessages(appended.Succeeded.Values, true);
                    }
                    else
                    {
                        client.DeleteMessages(appended.Succeeded.Values);
                        client.CommitDeletes();
                    }

                    Console.WriteLine($"After:  {client.ListMessages().Count} message(s) in '{folderName}'");
                }
                finally
                {
                    client.DeleteFolder(folderName);
                }
            }
        }
    }
}
