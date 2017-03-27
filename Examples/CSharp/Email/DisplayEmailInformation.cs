using System.IO;
using System;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.CSharp.Email
{
    class DisplayEmailInformation
    {
        public static void Run()
        {
            // ExStart:DisplayEmailInformation
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();

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
            // ExEnd:DisplayEmailInformation
        }
    }
}
