// Demonstrates the task-based SMTP client created from an asynchronous token provider.
//
// SmtpClient.CreateAsync builds the client from an IAsyncTokenProvider and returns it
// as IAsyncSmtpClient, the interface that holds the task-based API. The provider below
// obtains tokens from Microsoft Entra ID with MSAL; register an app with the delegated
// SMTP.Send permission, add http://localhost as a redirect URI, and fill in its client
// and tenant ids. SMTP AUTH must also be enabled for the mailbox.

using System;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Clients.Smtp.Models;
using Microsoft.Identity.Client;

namespace Aspose.Email.Examples.SMTP
{
    internal static class CreateAsyncSmtpClient
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            const string userName = "user@example.com";

            using (var tokenProvider = new MsalTokenProvider("<CLIENT_ID>", "<TENANT_ID>",
                new[] { "https://outlook.office.com/SMTP.Send" }))
            using (var cancellation = new CancellationTokenSource(TimeSpan.FromMinutes(3)))
            using (var client = await SmtpClient.CreateAsync("smtp.office365.com", userName,
                tokenProvider, 587, SecurityOptions.SSLExplicit, cancellation.Token))
            {
                var valid = await client.ValidateCredentialsAsync(token: cancellation.Token);
                Console.WriteLine($"Signed in: {valid}");

                await client.SendAsync(SmtpSend.Create()
                    .AddMessage(new MailMessage(userName, userName, "Sent with OAuth", "Body"))
                    .SetCancellationToken(cancellation.Token));

                Console.WriteLine($"Sent a message to {userName}.");
            }
        }

        // Signs the user in interactively in the system browser, then reuses the token
        // until shortly before it expires.
        private sealed class MsalTokenProvider : IAsyncTokenProvider
        {
            private readonly IPublicClientApplication _application;
            private readonly string[] _scopes;
            private AuthenticationResult _lastResult;

            public MsalTokenProvider(string clientId, string tenantId, string[] scopes)
            {
                _application = PublicClientApplicationBuilder
                    .Create(clientId)
                    .WithTenantId(tenantId)
                    .WithRedirectUri("http://localhost")
                    .Build();
                _scopes = scopes;
            }

            public async Task<OAuthToken> GetAccessTokenAsync(bool ignoreExistingToken = false,
                CancellationToken cancellationToken = default(CancellationToken))
            {
                var stillValid = _lastResult != null && _lastResult.ExpiresOn > DateTimeOffset.UtcNow.AddMinutes(5);

                if (ignoreExistingToken || !stillValid)
                {
                    _lastResult = await _application
                        .AcquireTokenInteractive(_scopes)
                        .WithUseEmbeddedWebView(false)
                        .ExecuteAsync(cancellationToken)
                        .ConfigureAwait(false);
                }

                return new OAuthToken(_lastResult.AccessToken);
            }

            public void Dispose()
            {
            }
        }
    }
}
