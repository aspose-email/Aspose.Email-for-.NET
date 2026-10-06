using System.Net;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Graph;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients.Pop3;
using Aspose.Email.Clients.Smtp;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using System;

namespace Aspose.Email.Examples
{

    public static class ClientBuilder
    {
        private static readonly IConfiguration ClientSettings;

        static ClientBuilder()
        {
            ClientSettings = new ConfigurationBuilder()
                .AddJsonFile(@"clientsettings.json")
                .Build();
        }

        public static ImapClient Imap(AuthType authType)
        {
            EmailClientSettings settings = new EmailClientSettings();
            ClientSettings.GetSection("Imap")
                .Bind(settings);

            switch (authType)
            {
                case AuthType.Basic:
                    return new ImapClient(settings.HostName, settings.Port, settings.UserName, settings.Password);
                case AuthType.ModernWithAppPermission:
                    throw new NotSupportedException(
                    "Access via application permissions is not supported. Delegated authentication only supported.");
                case AuthType.ModernWithDelegatedPermission:
                    return new ImapClient(settings.HostName, settings.Port,
                    settings.UserName, new TokenProvider(settings, authType));
                default:
                    throw new ArgumentOutOfRangeException(nameof(authType), authType, null);
            }
        }

        // True when clientsettings.json names an SMTP host. Examples that send mail check
        // this so that they stay runnable offline instead of hanging on a connect timeout.
        public static bool IsSmtpConfigured =>
            !string.IsNullOrWhiteSpace(ClientSettings.GetSection("Smtp")["HostName"]);

        public static SmtpClient Smtp(AuthType authType)
        {
            EmailClientSettings settings = new EmailClientSettings();
            ClientSettings.GetSection("Smtp")
                .Bind(settings);

            switch (authType)
            {
                case AuthType.Basic:
                    return new SmtpClient(settings.HostName, settings.Port, settings.UserName, settings.Password);
                case AuthType.ModernWithAppPermission:
                    throw new NotSupportedException(
                    "Access via application permissions is not supported. Delegated authentication only supported.");
                case AuthType.ModernWithDelegatedPermission:
                    return new SmtpClient(settings.HostName, settings.Port,
                    settings.UserName, new TokenProvider(settings, authType));
                default:
                    throw new ArgumentOutOfRangeException(nameof(authType), authType, null);
            }
        }

        public static Pop3Client Pop(AuthType authType)
        {
            EmailClientSettings settings = new EmailClientSettings();
            ClientSettings.GetSection("Pop")
                .Bind(settings);

            switch (authType)
            {
                case AuthType.Basic:
                    return new Pop3Client(settings.HostName, settings.Port, settings.UserName, settings.Password);
                case AuthType.ModernWithAppPermission:
                    throw new NotSupportedException(
                    "Access via application permissions is not supported. Delegated authentication only supported.");
                case AuthType.ModernWithDelegatedPermission:
                    return new Pop3Client(settings.HostName, settings.Port,
                    settings.UserName, new TokenProvider(settings, authType), SecurityOptions.None);
                default:
                    throw new ArgumentOutOfRangeException(nameof(authType), authType, null);

            }
        }

        public static IEWSClient Ews(AuthType authType)
        {
            EwsClientSettings settings = new EwsClientSettings();
            ClientSettings.GetSection("Ews")
                .Bind(settings);

            switch (authType)
            {
                case AuthType.Basic:
                    return EWSClient.GetEWSClient(settings.MailboxUri, settings.UserName, settings.Password);
                case AuthType.ModernWithAppPermission:
                case AuthType.ModernWithDelegatedPermission:
                    {
                        var tokenProvider = new TokenProvider(settings, authType);
                        NetworkCredential credentials =
                            new OAuthNetworkCredential(settings.UserName, tokenProvider.GetAccessToken()?.Token);
                        return EWSClient.GetEWSClient("https://outlook.office365.com/EWS/Exchange.asmx", credentials);
                    }
                default:
                    throw new ArgumentOutOfRangeException(nameof(authType), authType, null);
            }
        }

        // True when clientsettings.json carries the Graph application registration.
        // The Graph examples check this so that they explain what is missing instead of
        // failing with an authentication error.
        public static bool IsGraphConfigured =>
            !string.IsNullOrWhiteSpace(ClientSettings.GetSection("Graph")["ClientId"]) &&
            !string.IsNullOrWhiteSpace(ClientSettings.GetSection("Graph")["TenantId"]);

        // The mailbox the Graph examples act on. With application permissions the client
        // has no signed-in user, so the mailbox has to be named explicitly.
        public static string GraphMailboxId => ClientSettings.GetSection("Graph")["MailboxId"];

        public static IGraphClient Graph(AuthType authType)
        {
            var settings = GraphSettings();
            var client = GraphClient.GetClient(new TokenProvider(settings, RequireModern(authType)), settings.TenantId);
            Configure(client, settings);
            return client;
        }

        public static IGraphClientAsync GraphAsync(AuthType authType)
        {
            var settings = GraphSettings();
            var client = GraphClient.GetClientAsync(new TokenProvider(settings, RequireModern(authType)), settings.TenantId);
            Configure(client, settings);
            return client;
        }

        private static GraphClientSettings GraphSettings()
        {
            var settings = new GraphClientSettings();
            ClientSettings.GetSection("Graph").Bind(settings);
            return settings;
        }

        // Graph is OAuth-only; there is no basic-authentication path to fall back on.
        private static AuthType RequireModern(AuthType authType)
        {
            if (authType == AuthType.Basic)
                throw new NotSupportedException("Microsoft Graph supports OAuth 2.0 only.");

            return authType;
        }

        // IGraphClient and IGraphClientAsync declare these settings separately rather
        // than sharing a base interface, so each one is configured on its own.
        private static void Configure(IGraphClient client, GraphClientSettings settings)
        {
            if (!string.IsNullOrWhiteSpace(settings.EndPoint))
                client.EndPoint = settings.EndPoint;

            // With application permissions the token belongs to the app rather than to a
            // person, so the mailbox to work on is addressed by its id.
            if (!string.IsNullOrWhiteSpace(settings.MailboxId))
            {
                client.Resource = ResourceType.Users;
                client.ResourceId = settings.MailboxId;
            }
        }

        private static void Configure(IGraphClientAsync client, GraphClientSettings settings)
        {
            if (!string.IsNullOrWhiteSpace(settings.EndPoint))
                client.EndPoint = settings.EndPoint;

            if (!string.IsNullOrWhiteSpace(settings.MailboxId))
            {
                client.Resource = ResourceType.Users;
                client.ResourceId = settings.MailboxId;
            }
        }

        private class GraphClientSettings : IModernAuthSettings
        {
            public string UserName { get; set; }

            public string Password { get; set; }

            public string ClientId { get; set; }

            public string TenantId { get; set; }

            public string RedirectUri { get; set; }

            public string[] Scope { get; set; }

            public string ClientSecret { get; set; }

            public string EndPoint { get; set; }

            public string MailboxId { get; set; }
        }

        private interface IModernAuthSettings
        {
            string UserName { get; set; }

            string Password { get; set; }

            string ClientId { get; set; }

            string TenantId { get; set; }

            string RedirectUri { get; set; }

            string[] Scope { get; set; }

            string ClientSecret { get; set; }
        }

        private class EmailClientSettings : IModernAuthSettings
        {
            public string HostName { get; set; }

            public int Port { get; set; }

            public string UserName { get; set; }

            public string Password { get; set; }

            public string ClientId { get; set; }

            public string TenantId { get; set; }

            public string RedirectUri { get; set; }

            public string[] Scope { get; set; }
            public string ClientSecret { get; set; }
        }

        private class EwsClientSettings : IModernAuthSettings
        {
            public string MailboxUri { get; set; }

            public string Domain { get; set; }

            public string UserName { get; set; }

            public string Password { get; set; }

            public string ClientId { get; set; }

            public string TenantId { get; set; }

            public string RedirectUri { get; set; }

            public string[] Scope { get; set; }

            public string ClientSecret { get; set; }
        }

        private class TokenProvider : ITokenProvider
        {
            private readonly IModernAuthSettings _settings;
            private readonly AuthType _authType;
            private OAuthToken _token;

            public TokenProvider(IModernAuthSettings settings, AuthType authType)
            {
                _settings = settings;
                _authType = authType;
                _token = null;
            }

            public void Dispose()
            {
                throw new NotImplementedException();
            }

            public OAuthToken GetAccessToken()
            {
                try
                {
                    switch (_authType)
                    {
                        case AuthType.Basic:
                            throw new NotSupportedException();
                        case AuthType.ModernWithAppPermission:
                            {
                                var cca = ConfidentialClientApplicationBuilder
                                    .Create(_settings.ClientId)
                                    .WithClientSecret(_settings.ClientSecret)
                                    .WithTenantId(_settings.TenantId)
                                    .Build();
                                var result = cca.AcquireTokenForClient(_settings.Scope)
                                    .ExecuteAsync().GetAwaiter().GetResult()
                                    .AccessToken;
                                _token = new OAuthToken(result);
                                return _token;
                            }
                        case AuthType.ModernWithDelegatedPermission:
                            {
                                var pca = PublicClientApplicationBuilder
                                    .Create(_settings.ClientId)
                                    .WithTenantId(_settings.TenantId)
                                    .WithRedirectUri(_settings.RedirectUri)
                                    .Build();
                                var result = pca.AcquireTokenInteractive(_settings.Scope)
                                    .WithUseEmbeddedWebView(false)
                                    .ExecuteAsync().GetAwaiter().GetResult()
                                    .AccessToken;
                                _token = new OAuthToken(result);
                                return _token;
                            }
                        default:
                            throw new ArgumentOutOfRangeException();
                    }
                }
                catch (MsalException ex)
                {
                    Console.WriteLine($"Error acquiring access token: {ex}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex}");
                }

                return null;
            }

            public OAuthToken GetAccessToken(bool ignoreExistingToken)
            {
                return ignoreExistingToken ? GetAccessToken() : _token ?? GetAccessToken();
            }
        }
    }
}