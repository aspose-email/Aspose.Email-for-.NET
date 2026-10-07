// Demonstrates sending through an on-premises Exchange or other Windows-integrated SMTP
// server as the user the program runs as, without storing a password.
//
// UseDefaultCredentials makes the client authenticate with the current Windows logon
// (NTLM), so no user name or password is set. It only works against servers that accept
// NTLM - typically an internal Exchange receive connector - and only on Windows.

using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SendWithWindowsCredentials
    {
        public static void Run()
        {
            using (var client = new SmtpClient("exchange.corp.example.com", 587, SecurityOptions.SSLExplicit))
            {
                client.UseDefaultCredentials = true;
                client.AllowedAuthentication = SmtpKnownAuthenticationType.NTLM;

                var message = new MailMessage("me@corp.example.com", "colleague@corp.example.com",
                    "Build finished", "The nightly build finished without errors.");

                client.Send(message);
                Console.WriteLine($"Sent as {Environment.UserDomainName}\\{Environment.UserName} through {client.Host}.");
            }
        }
    }
}
