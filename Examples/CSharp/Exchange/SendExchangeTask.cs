using System.IO;
using System;
using Aspose.Email;
using Aspose.Email.Exchange;
using Aspose.Email.Mail;
using System.Net;
using System.Text;

namespace Aspose.Email.Examples.CSharp.Email.Exchange
{
    class SendExchangeTask
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Exchange();
            string dstEmail = dataDir + "Message.eml";
            
            // Create instance of ExchangeClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");


            // load task from .eml file
            EmlLoadOptions loadOptions = new EmlLoadOptions();

            loadOptions.PrefferedTextEncoding = Encoding.UTF8;
            loadOptions.PreserveTnefAttachments = true;

            // load task from .msg file
            MailMessage eml = MailMessage.Load(dstEmail, loadOptions);
            eml.From = "firstname.lastname@domain.com";
            eml.To.Clear();
            eml.To.Add(new Aspose.Email.Mail.MailAddress("firstname.lastname@domain.com"));

            client.Send(eml);

            Console.WriteLine(Environment.NewLine + "Task sent on Exchange Server successfully.");
        }
    }
}
