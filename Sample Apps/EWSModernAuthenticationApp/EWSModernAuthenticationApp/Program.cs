﻿using System;
using System.Net;
using System.Threading.Tasks;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Exchange.WebService;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;

namespace EWSModernAuthenticationApp
{
    // The following is the code sample that demonstrates making 
    // a modern-authenticated EWS request using application-only authentication.
    // Before running this sample, paste your username, app ID, tenant ID and client secret
    // in the proper sections of appsettings.json.
    class Program
    {
        static async Task Main(string[] args)
        {
            // Read access parameters (username, app Id, tenant Id and client secret) from appsettings.json 
            var accessParameters = ReadSettings("appsettings.json");

            //  Get an authentication token from a token server
            var accessToken = await GetAccessToken(accessParameters);

            // Using assess token to authenticate
            if (accessToken != null)
            {
                NetworkCredential credentials = new OAuthNetworkCredential(accessParameters.UserName, accessToken);

                using var client = EWSClient.GetEWSClient("https://outlook.office365.com/EWS/Exchange.asmx", credentials);
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
                UserName = appSettings.GetSection("credentials")["userName"],
                TenantId = appSettings.GetSection("credentials")["tenantId"],
                AppId = appSettings.GetSection("credentials")["appId"],
                ClientSecret = appSettings.GetSection("credentials")["clientSecret"]
            };
        }

        /// <summary>Get an authentication token from a token server.</summary>
        /// <param name="accessParameters">The access parameters.</param>
        /// <returns>The string that represents an authentication token.</returns>
        private static async Task<string> GetAccessToken(AccessParameters accessParameters)
        {
            var cca = ConfidentialClientApplicationBuilder
                .Create(accessParameters.AppId)
                .WithClientSecret(accessParameters.ClientSecret)
                .WithTenantId(accessParameters.TenantId)
                .Build();

            try
            {
                var result = await cca.AcquireTokenForClient(accessParameters.Scopes).ExecuteAsync();

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