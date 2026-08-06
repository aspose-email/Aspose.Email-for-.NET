// Demonstrates the X.500 address of a mail address. Messages sent inside an Exchange
// organisation identify the sender by a directory address rather than an SMTP one;
// that address lands in X500Address, leaving Address empty.

using System;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Email
{
    internal static class GetX500AddressFromMailAddress
    {
        private const string ExchangeAddress =
            "/O=EXCHANGELABS/OU=EXCHANGE ADMINISTRATIVE GROUP (FYDIBOHF23SPDLT)/CN=RECIPIENTS/CN=sender01";

        public static void Run()
        {
            // Build an Exchange-style message: the "EX" address type is what marks the
            // sender address as an X.500 one.
            var msg = new MapiMessage
            {
                Subject = "Message from inside the organisation",
                Body = "Sent through Exchange."
            };
            msg.SetProperty(KnownPropertyList.SenderAddressType, "EX");
            msg.SetProperty(KnownPropertyList.SenderEmailAddress, ExchangeAddress);
            msg.SetProperty(KnownPropertyList.SenderName, "Alice Johnson");
            msg.Recipients.Add("bob@example.com", "SMTP", "Bob Smith", MapiRecipientType.MAPI_TO);

            var msgPath = Data.Out/"GetX500AddressFromMailAddress_out.msg";
            msg.Save(msgPath);

            var mailMessage = MailMessage.Load(msgPath, new MsgLoadOptions());

            Print("From", mailMessage.From);
            foreach (var to in mailMessage.To)
                Print("To", to);

            Console.WriteLine($"\nSaved to {msgPath}");
        }

        private static void Print(string label, MailAddress address)
        {
            if (address == null)
                return;

            Console.WriteLine($"{label}: {address.DisplayName}");
            Console.WriteLine($"  SMTP address:  {Show(address.Address)}");
            Console.WriteLine($"  X.500 address: {Show(address.X500Address)}");
        }

        private static string Show(string value) => string.IsNullOrEmpty(value) ? "(none)" : value;
    }
}
