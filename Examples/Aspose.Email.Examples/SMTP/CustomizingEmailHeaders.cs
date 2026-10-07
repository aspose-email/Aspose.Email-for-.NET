// Demonstrates how to set the standard address and date headers of a message and add a
// custom one, then save it as an Outlook .msg file.
//
// Custom headers - by convention starting with "X-" - travel with the message and can be
// read by the receiving side, for example to tag mail sent by your application.
// CustomizingEmailHeader sends a message built the same way.

using System;

namespace Aspose.Email.Examples.SMTP
{
    internal static class CustomizingEmailHeaders
    {
        public static void Run()
        {
            var message = new MailMessage
            {
                From = "sender@example.com",
                Subject = "Test mail",
                Date = new DateTime(2026, 3, 6),
                XMailer = "Aspose.Email"
            };

            message.To.Add("receiver1@example.com");
            message.CC.Add("receiver2@example.com");
            message.Bcc.Add("receiver3@example.com");
            message.ReplyToList.Add("reply@example.com");
            message.Headers.Add("X-Secret-Header", "mystery");

            var outputPath = Data.Out/"CustomizingEmailHeaders_out.msg";
            message.Save(outputPath, SaveOptions.DefaultMsgUnicode);
            Console.WriteLine($"Saved to {outputPath}");

            var loaded = MailMessage.Load(outputPath);
            Console.WriteLine("\nRead back:");
            Console.WriteLine($"  subject:         {loaded.Subject}");
            Console.WriteLine($"  reply to:        {loaded.ReplyToList}");
            Console.WriteLine($"  X-Secret-Header: {loaded.Headers["X-Secret-Header"]}");
        }
    }
}
