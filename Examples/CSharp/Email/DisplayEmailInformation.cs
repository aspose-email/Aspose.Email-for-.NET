using System.IO;
using System;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;

namespace Aspose.Email.Examples.CSharp.Email
{
    class DisplayEmailInformation
    {
        public static void Run()
        {
            // ExStart:DisplayEmailInformation
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();
            string dstEmail = dataDir + "test.eml";

            // Create MailMessage instance by loading an Eml file
            MailMessage message = MailMessage.Load(dataDir + "test.eml", new EmlLoadOptions());

            // Gets the sender info, recipient info, Subject, htmlbody and textbody
            Console.Write("From:");
            Console.WriteLine(message.From);
            Console.Write("To:");
            Console.WriteLine(message.To);
            Console.Write("Subject:");
            Console.WriteLine(message.Subject);
            Console.WriteLine("HtmlBody:");
            Console.WriteLine(message.HtmlBody);
            Console.WriteLine("TextBody");
            Console.WriteLine(message.Body);
            Console.WriteLine(Environment.NewLine + "Displayed email information from " + dstEmail);
            // ExEnd:DisplayEmailInformation
        }
    }
}
