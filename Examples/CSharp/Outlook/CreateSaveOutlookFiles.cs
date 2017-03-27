using System.IO;
using System;
using Aspose.Email;
using Aspose.Email.Storage.Pst;
using Aspose.Email.Mapi;
using Aspose.Email.Mime;

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class CreateSaveOutlookFiles
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Outlook();
            string dst = dataDir + "message.msg";

            // Create an instance of MailMessage class
            MailMessage mailMsg = new MailMessage();

            // Set FROM field of the message
            mailMsg.From = "from@domain.com";

            // Set TO field of the message
            mailMsg.To.Add("to@domain.com");

            // Set SUBJECT of the message
            mailMsg.Subject = "creating an outlook message file";

            // Set BODY of the message
            mailMsg.Body = "This message is created by Aspose.Email";

            // Create an instance of MapiMessage class and pass MailMessage as argument
            MapiMessage outlookMsg = MapiMessage.FromMailMessage(mailMsg);

            // Save the message (msg) file
            outlookMsg.Save(dst);

            Console.WriteLine(Environment.NewLine + "MSG saved successfully at " + dst);
        }
    }
}
