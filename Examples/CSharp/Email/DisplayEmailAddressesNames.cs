using Aspose.Email.Mime;

namespace Aspose.Email.Examples.CSharp.Email
{
    class DisplayEmailAddressesNames
    {
        public static void Run()
        {
            // ExStart:DisplayEmailAddressesNames
            // Declare message as MailMessage instance
            MailMessage message = new MailMessage("Sender <sender@from.com>", "Recipient <recipient@to.com>");
 
            // Specify the recipients mail addresses
            message.To.Add("receiver1@receiver.com");
            message.To.Add("receiver2@receiver.com");
            message.To.Add("receiver3@receiver.com");
            message.To.Add("receiver4@receiver.com");

            // Specifying CC and BCC Addresses
            message.CC.Add("CC1@receiver.com");
            message.CC.Add("CC2@receiver.com");
            message.Bcc.Add("Bcc1@receiver.com");
            message.Bcc.Add("Bcc2@receiver.com");
            // ExEnd:DisplayEmailAddressesNames
        }
    }
}
