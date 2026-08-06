// Demonstrates the two sets of load options that apply when reading an mbox:
// MboxLoadOptions configures the storage, EmlLoadOptions configures each message.

using System;
using System.Text;
using Aspose.Email.Storage.Mbox;

namespace Aspose.Email.Examples.MBOX
{
    internal static class ReadMboxWithLoadOptions
    {
        public static void Run()
        {
            // PreferredTextEncoding is used for messages that declare no encoding of
            // their own, which is common in older mbox files.
            var mboxOptions = new MboxLoadOptions
            {
                PreferredTextEncoding = Encoding.UTF8,
                LeaveOpen = false
            };

            // Per-message options: keep any winmail.dat attachment as-is instead of
            // unpacking it into ordinary attachments.
            var emlOptions = new EmlLoadOptions { PreserveTnefAttachments = true };

            using (var reader = new MboxrdStorageReader(Data.Mbox/"ExampleMbox.mbox", mboxOptions))
            {
                MailMessage message;
                while ((message = reader.ReadNextMessage(emlOptions)) != null)
                {
                    using (message)
                    {
                        Console.WriteLine($"Subject:      {message.Subject}");
                        Console.WriteLine($"Originally TNEF: {message.OriginalIsTnef}");
                        Console.WriteLine($"Attachments:  {message.Attachments.Count}");
                        Console.WriteLine("-----------------------------------");
                    }
                }
            }
        }
    }
}
