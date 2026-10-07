// Demonstrates delivering messages through a pickup directory instead of the network.
//
// With DeliveryMethod set to SpecifiedPickupDirectory, Send writes each message as a
// file into PickupDirectoryLocation (an absolute path) and returns at once. A local mail
// server that watches that folder - IIS SMTP, Exchange, hMailServer and others - picks
// the files up and delivers them. PickupDirectoryFromIis uses the IIS SMTP folder.
// No server is involved here, so this example runs offline.

using System;
using System.IO;
using System.Linq;
using Aspose.Email.Clients.Smtp;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SaveToPickupDirectory
    {
        public static void Run()
        {
            var pickupDir = Data.OutSub("SmtpPickup");

            using (var client = new SmtpClient())
            {
                client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                client.PickupDirectoryLocation = pickupDir;

                for (var i = 1; i <= 3; i++)
                {
                    var message = new MailMessage("sender@example.com", "receiver@example.com",
                        $"Order confirmation {1000 + i}", "Thank you for your order.");
                    client.Send(message);
                }
            }

            // The file names are generated, so order by the time the files were written.
            var files = Directory.GetFiles(pickupDir).OrderBy(File.GetLastWriteTimeUtc).ToList();
            Console.WriteLine($"{files.Count} file(s) written to {pickupDir}:");

            foreach (var file in files)
            {
                var message = MailMessage.Load(file);
                Console.WriteLine($"  {Path.GetFileName(file)}");
                Console.WriteLine($"    {message.From} -> {message.To}: {message.Subject}");
            }
        }
    }
}
