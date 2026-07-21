// Demonstrates how to load email messages in various formats (EML, HTML, MHTML, MSG)
// using format-specific load options, including custom encoding and TNEF preservation.

using System;
using System.Text;

namespace Aspose.Email.Examples.Email
{
    internal static class LoadMessageWithLoadOptions
    {
        public static void Run()
        {
            // Passing the matching load options tells Aspose.Email which parser to use,
            // instead of letting it detect the format from the content.
            Report("EML  ", MailMessage.Load(Data.Email/"Message.eml", new EmlLoadOptions()));
            Report("HTML ", MailMessage.Load(Data.Email/"description.html", new HtmlLoadOptions()));
            Report("MHTML", MailMessage.Load(Data.Email/"Message.mhtml", new MhtmlLoadOptions()));
            Report("MSG  ", MailMessage.Load(Data.Email/"Message.msg", new MsgLoadOptions()));

            // PreferredTextEncoding is used when the message declares no encoding of its
            // own; PreserveTnefAttachments keeps TNEF parts as they are.
            var emlLoadOptions = new EmlLoadOptions
            {
                PreferredTextEncoding = Encoding.UTF8,
                PreserveTnefAttachments = true
            };
            Report("EML (UTF-8, TNEF preserved)", MailMessage.Load(Data.Email/"Message.eml", emlLoadOptions));

            // ShouldAddPlainTextView builds a text alternative from the HTML, and
            // PathToResources tells the loader where the images referenced by it live.
            var htmlLoadOptions = new HtmlLoadOptions
            {
                PreferredTextEncoding = Encoding.UTF8,
                ShouldAddPlainTextView = true,
                PathToResources = Data.Email
            };
            Report("HTML (UTF-8, plain text view)", MailMessage.Load(Data.Email/"description.html", htmlLoadOptions));
        }

        private static void Report(string label, MailMessage message)
        {
            var subject = string.IsNullOrEmpty(message.Subject) ? "(no subject)" : message.Subject;
            Console.WriteLine($"{label,-30} -> {subject} ({message.Attachments.Count} attachment(s), " +
                              $"{message.AlternateViews.Count} alternate view(s))");
        }
    }
}
