using System;
using System.Threading.Tasks;
using Aspose.Email;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients.Smtp;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;

namespace EWSModernAuthenticationImapSmtp
{
	// The following is the code sample that demonstrates making 
	// modern-authenticated IMAP, SMTP requests using delegated authentication.
	// Before running this sample, paste your Office365 username, client ID and tenant ID 
	// in the proper sections of appsettings.json.
    class Program
    {
        static async Task Main(string[] args)
        {
            // Read access parameters (client Id, tenant Id and redirect URI) from appsettings.json 
            var accessParameters = ReadSettings("appsettings.json");

            //  Get an authentication token from a token server
            var accessToken = await GetAccessToken(accessParameters);

            if (accessToken != null)
            {
                // Using assess token to authenticate and list messages via IMAP.
                ListMessagesViaImap(accessParameters.UserName, accessToken);

                // Using assess token to authenticate and send message via SMTP.
                SendMessageViaSmtp(accessParameters.UserName, accessToken);
            }
        }

        /// <summary>
        /// Lists the Inbox messages via IMAP.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="accessToken">The access token.</param>
        private static void ListMessagesViaImap(string userName, string accessToken)
        {
            Console.WriteLine("----List Inbox messages");

            using var imapClient = new ImapClient(
                "outlook.office365.com",
                993,
                userName,
                accessToken,
                true)
            {
                SecurityOptions = SecurityOptions.SSLImplicit
            };

            imapClient.SelectFolder(ImapFolderInfo.InBox);
            var messageInfoCollection = imapClient.ListMessages();

            foreach (ImapMessageInfo imapMessageInfo in messageInfoCollection)
            {
                Console.WriteLine(imapMessageInfo.Subject);
            }

            Console.WriteLine();
        }

        /// <summary>
        /// Sends the message via SMTP.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="accessToken">The access token.</param>
        private static void SendMessageViaSmtp(string userName, string accessToken)
        {
            Console.WriteLine("----Send message");

            using var smtpClient = new SmtpClient(
                "smtp.office365.com",
                587,
                userName,
                accessToken,
                true)
            {
                SecurityOptions = SecurityOptions.SSLAuto
            };

            var eml = new MailMessage(userName, userName, "test message", "this is a test message");
            smtpClient.Send(eml);

            Console.WriteLine();
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
                UserName = appSettings.GetSection("credentials")["userName"],
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
