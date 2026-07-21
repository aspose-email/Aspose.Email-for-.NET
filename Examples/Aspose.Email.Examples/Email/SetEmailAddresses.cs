// Demonstrates various ways to set sender, To, CC, and BCC addresses on a MailMessage,
// including string assignment, Add, and MailAddress/EmailAddress objects.

using System;
using Aspose.Email.PersonalInfo;

namespace Aspose.Email.Examples.Email
{
    internal static class SetEmailAddresses
    {
        public static void Run()
        {
            // Addresses can be given in "Display Name <address>" form.
            var eml = new MailMessage("Kurt Bracy <KurtBracy@from.com>", "Mary Mack <MaryMack@to.com>")
            {
                Subject = "Sample eml",
                Body = "It is a sample eml"
            };

            // Assigning a string replaces the whole collection, so include the existing
            // recipients to add to them rather than overwrite them.
            eml.To = $"{eml.To}, JeanJetton@to.com, LoriRocha@to.com";

            // Add appends a single address.
            eml.To.Add("MonroeTay@to.com");
            eml.To.Add("JimmieCamacho@to.com");

            eml.CC = "RoscoeCummings@cc.com, MarkMohr@cc.com";

            // A MailAddress or an EmailAddress object works just as well as a string.
            eml.Bcc = new MailAddress("RudolphHancock@bcc.com", "Rudolph Hancock");
            eml.Bcc.Add(new EmailAddress("AdamFritts@bcc.com", "Adam Fritts"));

            Console.WriteLine($"From: {eml.From}");
            Console.WriteLine($"To  ({eml.To.Count}): {eml.To}");
            Console.WriteLine($"CC  ({eml.CC.Count}): {eml.CC}");
            Console.WriteLine($"Bcc ({eml.Bcc.Count}): {eml.Bcc}");
        }
    }
}
