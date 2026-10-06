// Demonstrates the two ways ImapClient can sign in with OAuth 2.0 (SASL XOAUTH2)
// instead of a password.
//
// - Access token: pass it where the password would go and set useOAuth to true. The
//   client uses the token as is, so it stops working once the token expires.
// - Token provider: the client asks an ITokenProvider for a token whenever it needs
//   one, so the token can be refreshed without recreating the client.
//   TokenProvider.Google and TokenProvider.Outlook exchange a refresh token for access
//   tokens; ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission) shows a
//   provider built on MSAL instead.

using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class ConnectUsingOAuth
    {
        public static void Run()
        {
            const string host = "imap.gmail.com";
            const int port = 993;
            const string userName = "user@gmail.com";

            // 1. An access token obtained elsewhere.
            const string accessToken = "<ACCESS_TOKEN>";

            using (var client = new ImapClient(host, port, userName, accessToken, true, SecurityOptions.SSLImplicit))
            {
                client.SelectFolder(ImapFolderInfo.InBox);
                Console.WriteLine($"Access token:   {client.CurrentFolder.TotalMessageCount} message(s) in the Inbox");
            }

            // 2. A provider that turns a refresh token into access tokens on demand.
            var tokenProvider = TokenProvider.Google.GetInstance("<CLIENT_ID>", "<CLIENT_SECRET>", "<REFRESH_TOKEN>");

            using (var client = new ImapClient(host, port, userName, tokenProvider, SecurityOptions.SSLImplicit))
            {
                client.SelectFolder(ImapFolderInfo.InBox);
                Console.WriteLine($"Token provider: {client.CurrentFolder.TotalMessageCount} message(s) in the Inbox");
            }
        }
    }
}
