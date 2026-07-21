using System;
using System.Text;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;


namespace Aspose.Email.Examples.EWS
{
    class SendExchangeTask
    {
        public static void Run()
        {

            try
            {

                // The path to the File directory.

                string dstEmail = Data.Out + "Message.eml";

                // Create instance of ExchangeClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // load task from .eml file
                EmlLoadOptions loadOptions = new EmlLoadOptions();

                loadOptions.PreferredTextEncoding = Encoding.UTF8;
                loadOptions.PreserveTnefAttachments = true;

                // load task from .msg file
                MailMessage eml = MailMessage.Load(dstEmail, loadOptions);
                eml.From = "firstname.lastname@domain.com";
                eml.To.Clear();
                eml.To.Add(new MailAddress("firstname.lastname@domain.com"));
                client.Send(eml);
                Console.WriteLine(Environment.NewLine + "Task sent on Exchange Server successfully.");
            }
            catch (Exception ex)
            {

                Console.Write(ex.Message);
            }
        }
    }
}
