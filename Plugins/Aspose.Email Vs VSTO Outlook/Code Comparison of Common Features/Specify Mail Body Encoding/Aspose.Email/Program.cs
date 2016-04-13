using Aspose.Email.Mail;
using System;
using System.Collections.Generic;
using System.Text;

namespace Aspose.Email
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create an Instance of MailMessage class
            MailMessage message = new MailMessage();

            //From field
            message.From = "sender@sender.com";

            //To field
            message.To.Add("receiver@receiver.com");

            //Specify HtmlBody
            message.HtmlBody = "<html><body>This is the Html body</body></html>";

            //Specify BodyEncoding as ASCII
            message.BodyEncoding = Encoding.ASCII;

            //Create an instance of SmtpClient Class
            SmtpClient client = new SmtpClient();

            //Specify your mailing host server
            client.Host = "smtp.server.com";

            //Specify your mail user name
            client.Username = "Username";

            //Specify your mail password
            client.Password = "Password";

            //Specify your Port #
            client.Port = 25;

            try
            {
                //Client.Send will send this message
                client.Send(message);
                //Display ‘Message Sent’, only if message sent successfully
                Console.WriteLine("Message sent");
            }

            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }

            Console.WriteLine("Press enter to quit");
            Console.Read();
        }
    }
}
