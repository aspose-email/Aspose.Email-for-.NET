using System;
using System.Net;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    class SendTaskRequestUsingIEWSClient
    {
        public static void Run()
        {
            try
            {
                // Create instance of ExchangeClient class by giving credentials
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                MsgLoadOptions options = new MsgLoadOptions();
                options.PreserveTnefAttachments = true;

                // load task from .msg file
                MailMessage eml = MailMessage.Load(Data.Ews + "task.msg", options);
                eml.From = "firstname.lastname@domain.com";
                eml.To.Clear();
                eml.To.Add(new MailAddress("firstname.lastname@domain.com"));
                client.Send(eml);
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }
        }
    }
}