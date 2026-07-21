// Demonstrates how to create a new email with attachments, then save it as both EML and MSG.

using System;

namespace Aspose.Email.Examples.MAPI
{
    internal static class CreateEmlFileAndConvertToMsg
    {
        public static void Run()
        {
            var message = new MailMessage("from@domain.com", "to@domain.com")
            {
                Subject = "subject of email",
                HtmlBody = "<b>Eml to msg conversion using Aspose.Email</b>" +
                    "<br><hr><br><font color=blue>This is a test eml file which will be converted to msg format.</font>"
            };

            message.Attachments.Add(new Attachment(Data.Mapi/"attachment_1.doc"));
            message.Attachments.Add(new Attachment(Data.Mapi/"download.png"));

            // The same in-memory message is written out in both formats - no conversion
            // step is needed between them.
            var emlPath = Data.Out/"CreatEMLFileAndConvertToMSG_out.eml";
            var msgPath = Data.Out/"CreatEMLFileAndConvertToMSG_out.msg";
            message.Save(emlPath, SaveOptions.DefaultEml);
            message.Save(msgPath, SaveOptions.DefaultMsgUnicode);

            Console.WriteLine($"Message with {message.Attachments.Count} attachment(s):");
            Console.WriteLine($"  {emlPath}");
            Console.WriteLine($"  {msgPath}");
        }
    }
}
