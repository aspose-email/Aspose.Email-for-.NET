using System.IO;
using System;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;
using Aspose.Email.Clients.Pop3;
using Aspose.Email;
using Aspose.Email.Clients.Imap;
using System.Configuration;
using System.Data;
using Aspose.Email.Bounce;

namespace Aspose.Email.Examples.CSharp.Email
{
    class ProcessBouncedMsgs
    {
        public static void Run()
        {
            // The path to the File directory.
            string dataDir = RunExamples.GetDataDir_Email();
            string dstEmail = dataDir + "test.eml";

            string fileName = dstEmail;
            MailMessage mail = MailMessage.Load(fileName);
            BounceResult result = mail.CheckBounced();
            Console.WriteLine(fileName);
            Console.WriteLine("IsBounced : " + result.IsBounced);
            Console.WriteLine("Action : " + result.Action);
            Console.WriteLine("Recipient : " + result.Recipient);            
            Console.WriteLine(Environment.NewLine + "Bounce information displayed successfully.");
        }
    }
}
