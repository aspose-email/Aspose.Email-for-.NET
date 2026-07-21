// Demonstrates how to read every message out of a Thunderbird mbox storage and save
// each one to disk in both EML and MSG format.

using System;
using System.IO;
using Aspose.Email.Storage.Mbox;

namespace Aspose.Email.Examples.MBOX
{
    internal static class ReadMessagesFromThunderbird
    {
        public static void Run()
        {
            var outputDir = Data.OutSub("Thunderbird");

            using (var stream = new FileStream(Data.Mbox/"ExampleMbox.mbox", FileMode.Open, FileAccess.Read))
            using (var reader = new MboxrdStorageReader(stream, new MboxLoadOptions()))
            {
                var count = 0;

                MailMessage message;
                while ((message = reader.ReadNextMessage()) != null)
                {
                    using (message)
                    {
                        Console.WriteLine($"Subject: {message.Subject}");

                        // A subject can contain characters that are not allowed in a file
                        // name, so sanitise it before saving under that name.
                        var safeName = ToFileName(message.Subject, count);

                        message.Save(outputDir/(safeName + ".eml"), SaveOptions.DefaultEml);
                        message.Save(outputDir/(safeName + ".msg"), SaveOptions.DefaultMsgUnicode);
                        count++;
                    }
                }

                Console.WriteLine($"\nSaved {count} message(s) as EML and MSG to {outputDir}");
            }
        }

        private static string ToFileName(string subject, int index)
        {
            if (string.IsNullOrWhiteSpace(subject))
                return $"message-{index}";

            foreach (var invalid in Path.GetInvalidFileNameChars())
                subject = subject.Replace(invalid, ' ');

            return subject.Trim();
        }
    }
}
