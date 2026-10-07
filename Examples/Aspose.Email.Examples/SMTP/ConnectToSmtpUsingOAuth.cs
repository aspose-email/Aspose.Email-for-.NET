// Demonstrates the two ways SmtpClient can sign in with OAuth 2.0 (SASL XOAUTH2)
// instead of a password.
//
// - Access token: pass it where the password would go and set useOAuth to true. The
//   client uses the token as is, so it stops working once the token expires.
// - Token provider: the client asks an ITokenProvider for a token whenever it needs
//   one, so the token can be refreshed without recreating the client.
//   TokenProvider.Google and TokenProvider.Outlook exchange a refresh token for access
//   tokens; ClientBuilder.Smtp(AuthType.ModernWithDelegatedPermission) shows a provider
//   built on MSAL instead.

using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;

namespace Aspose.Email.Examples.SMTP
{
    internal static class ConnectToSmtpUsingOAuth
    {
        public static void Run()
        {
            const string host = "smtp.gmail.com";
            const int port = 587;
            const string userName = "user@gmail.com";

            // 1. An access token obtained elsewhere.
            const string accessToken = "<ACCESS_TOKEN>";

            using (var client = new SmtpClient(host, port, userName, accessToken, true, SecurityOptions.SSLExplicit))
            {
                client.Send(new MailMessage(userName, userName, "Sent with an access token", "Body"));
                Console.WriteLine("Sent with an access token.");
            }

            // 2. A provider that turns a refresh token into access tokens on demand.
            var tokenProvider = TokenProvider.Google.GetInstance("<CLIENT_ID>", "<CLIENT_SECRET>", "<REFRESH_TOKEN>");

            using (var client = new SmtpClient(host, port, userName, tokenProvider, SecurityOptions.SSLExplicit))
            {
                client.Send(new MailMessage(userName, userName, "Sent with a token provider", "Body"));
                Console.WriteLine("Sent with a token provider.");
            }
        }
    }
}
