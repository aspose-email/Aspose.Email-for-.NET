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

            //Specify the recipients’ mail addresses
            message.To.Add("receiver1@receiver.com");
            message.To.Add("receiver2@receiver.com");
            message.To.Add("receiver3@receiver.com");

            message.CC.Add("CC1@receiver.com");
            message.CC.Add("CC2@receiver.com");

            message.Bcc.Add("Bcc1@receiver.com");
            message.Bcc.Add("Bcc2@receiver.com");

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
