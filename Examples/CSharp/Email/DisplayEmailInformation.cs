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
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Email();
            string dstEmail = dataDir + "test.eml";

            MailMessage message;

            //Create MailMessage instance by loading an Eml file
            message = MailMessage.Load(dataDir + "test.eml", new EmlLoadOptions());

            System.Console.Write("From:");

            //Gets the sender info
            System.Console.WriteLine(message.From);
            System.Console.Write("To:");

            //Gets the recipient info
            System.Console.WriteLine(message.To);
            System.Console.Write("Subject:");

            //Gets the subject
            System.Console.WriteLine(message.Subject);
            System.Console.WriteLine("HtmlBody:");

            //Gets the htmlbody 
            System.Console.WriteLine(message.HtmlBody);
            System.Console.WriteLine("TextBody");

            //Gets the textbody
            System.Console.WriteLine(message.Body);

            Console.WriteLine(Environment.NewLine + "Displayed email information from " + dstEmail);
        }
    }
}
