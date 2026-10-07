// Demonstrates keeping SMTP settings in a configuration file instead of in code.
//
// SmtpClient has no constructor that reads a configuration file, so read the settings
// with the configuration library of your choice and pass them to the client. This
// example reads the "Smtp" section of clientsettings.json with
// Microsoft.Extensions.Configuration - the same file the other examples use through
// ClientBuilder.

using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using Microsoft.Extensions.Configuration;

namespace Aspose.Email.Examples.SMTP
{
    internal static class LoadSmtpConfigFile
    {
        public static void Run()
        {
            var smtp = new ConfigurationBuilder()
                .AddJsonFile("clientsettings.json")
                .Build()
                .GetSection("Smtp");

            var host = smtp["HostName"];
            if (string.IsNullOrWhiteSpace(host))
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            var port = int.Parse(smtp["Port"] ?? "587");
            var userName = smtp["UserName"];

            Console.WriteLine($"Settings read from clientsettings.json: {userName} at {host}:{port}");

            using (var client = new SmtpClient(host, port, userName, smtp["Password"], SecurityOptions.Auto))
            {
                client.Send(new MailMessage(userName, userName, "Settings from a config file", "Body"));
                Console.WriteLine($"Sent to {userName}.");
            }
        }
    }
}
