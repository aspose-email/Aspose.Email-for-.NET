// Demonstrates how to create a MailMessage with sender, recipients, CC, and an HTML body,
// then save it in EML, EMLX, MSG, and MHTML formats.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class CreateNewMailMessage
    {
        public static void Run()
        {
            var message = new MailMessage
            {
                Subject = "New message created by Aspose.Email for .NET",
                HtmlBody = "<b>This line is in bold.</b> <br/> <br/>" +
                           "<font color=blue>This line is in blue color</font>",
                From = new MailAddress("from@domain.com", "Sender Name", false)
            };

            message.To.Add(new MailAddress("to1@domain.com", "Recipient 1", false));
            message.To.Add(new MailAddress("to2@domain.com", "Recipient 2", false));
            message.CC.Add(new MailAddress("cc1@domain.com", "Recipient 3", false));
            message.CC.Add(new MailAddress("cc2@domain.com", "Recipient 4", false));

            // EMLX has no dedicated Default property, so its options are built by type.
            message.Save(Data.Out/"CreateNewMailMessage_out.eml", SaveOptions.DefaultEml);
            message.Save(Data.Out/"CreateNewMailMessage_out.emlx", SaveOptions.CreateSaveOptions(MailMessageSaveType.EmlxFormat));
            message.Save(Data.Out/"CreateNewMailMessage_out.msg", SaveOptions.DefaultMsgUnicode);
            message.Save(Data.Out/"CreateNewMailMessage_out.mhtml", SaveOptions.DefaultMhtml);

            Console.WriteLine($"Subject: {message.Subject}");
            Console.WriteLine($"To: {message.To.Count} recipient(s), CC: {message.CC.Count}");
            Console.WriteLine($"Saved as EML, EMLX, MSG and MHTML in {Data.Out}");
        }
    }
}
