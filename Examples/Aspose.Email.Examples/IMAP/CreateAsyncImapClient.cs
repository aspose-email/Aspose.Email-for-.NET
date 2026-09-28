// Demonstrates the task-based IMAP client created from an asynchronous token provider.
//
// ImapClient.CreateAsync builds the client from an IAsyncTokenProvider and returns it
// as IAsyncImapClient, the interface that holds the task-based API. Every method of it
// accepts an optional CancellationToken. The provider below obtains tokens from
// Microsoft Entra ID with MSAL; register an app with the IMAP.AccessAsUser.All
// permission and fill in its client and tenant ids.

using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using Microsoft.Identity.Client;

namespace Aspose.Email.Examples.IMAP
{
    internal static class CreateAsyncImapClient
    {
        public static void Run()
        {
            RunAsync().GetAwaiter().GetResult();
        }

        private static async Task RunAsync()
        {
            using (var tokenProvider = new MsalTokenProvider("<CLIENT_ID>", "<TENANT_ID>",
                new[] { "https://outlook.office.com/IMAP.AccessAsUser.All" }))
            using (var cancellation = new CancellationTokenSource(TimeSpan.FromMinutes(3)))
            {
                var client = await ImapClient.CreateAsync("outlook.office365.com", "user@example.com",
                    tokenProvider, 993, SecurityOptions.SSLImplicit, cancellation.Token);

                try
                {
                    var valid = await client.ValidateCredentialsAsync(token: cancellation.Token);
                    Console.WriteLine($"Signed in: {valid}");

                    var folders = await client.ListFoldersAsync(token: cancellation.Token);
                    Console.WriteLine($"{folders.Count} folder(s) in the mailbox.");

                    await client.SelectFolderAsync(ImapFolderInfo.InBox, token: cancellation.Token);
                    var messages = await client.ListMessagesAsync(ImapFolderInfo.InBox, token: cancellation.Token);

                    Console.WriteLine($"{messages.Count} message(s) in the Inbox, newest first:");
                    foreach (var info in messages.OrderByDescending(info => info.InternalDate).Take(5))
                        Console.WriteLine($"  {info.InternalDate:g}  {info.Subject}");
                }
                finally
                {
                    (client as IDisposable)?.Dispose();
                }
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
