using System;
using System.Net;
using System.Threading.Tasks;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Exchange.WebService;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;

namespace EWSModernAuthenticationDelegated
{
	// The following is the code sample that demonstrates making 
	// a modern-authenticated EWS request using delegated authentication.
	// Before running this sample, paste your client ID and tenant ID 
	// in the proper sections of appsettings.json.
    class Program
    {
        static async Task Main(string[] args)
        {
            // Read access parameters (client Id, tenant Id and redirect URI) from appsettings.json 
            var accessParameters = ReadSettings("appsettings.json");

            //  Get an authentication token from a token server
            var accessToken = await GetAccessToken(accessParameters);

            // Using assess token to authenticate
            if (accessToken != null)
            {
                NetworkCredential credentials = new OAuthNetworkCredential(accessToken);

                using var client =
                    EWSClient.GetEWSClient("https://outlook.office365.com/EWS/Exchange.asmx", credentials);
                
                var messageInfos = client.ListMessages("Inbox");

                foreach (var messageInfo in messageInfos)
                {
                    Console.WriteLine(messageInfo.Subject);
                }
            }
        }

        /// <summary>Reads access parameters from settings file.</summary>
        /// <param name="settingsFileName">Name of the settings file.</param>
        /// <returns><see cref="AccessParameters"/></returns>
        private static AccessParameters ReadSettings(string settingsFileName)
        {
            IConfiguration appSettings = new ConfigurationBuilder()
                .AddJsonFile(settingsFileName)
                .Build();

            return new AccessParameters()
            {
                TenantId = appSettings.GetSection("credentials")["tenantId"],
                ClientId = appSettings.GetSection("credentials")["clientId"],
                RedirectUri = appSettings.GetSection("redirectUri").Value
            };
        }

        /// <summary>Get an authentication token from a token server.</summary>
        /// <param name="accessParameters">The access parameters.</param>
        /// <returns>The string that represents an authentication token.</returns>
        private static async Task<string> GetAccessToken(AccessParameters accessParameters)
        {
            var pcaOptions = new PublicClientApplicationOptions
            {
                ClientId = accessParameters.ClientId,
                TenantId = accessParameters.TenantId,
                RedirectUri = accessParameters.RedirectUri
            };

            var pca = PublicClientApplicationBuilder
                .CreateWithApplicationOptions(pcaOptions).Build();

            try
            {
                var result = await pca.AcquireTokenInteractive(accessParameters.Scopes)
                    .WithUseEmbeddedWebView(false)
                    .ExecuteAsync();
                
                return result.AccessToken;
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
    }
}
