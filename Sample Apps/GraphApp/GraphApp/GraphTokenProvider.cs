using Microsoft.Identity.Client;
using Aspose.Email.Clients;
using Microsoft.Identity.Web;

namespace GraphApp;

public class GraphTokenProvider : ITokenProvider
{
    private readonly IConfidentialClientApplication _app;
    private readonly string[] _scopes;
    private string? _token;
    
    public GraphTokenProvider(AuthenticationConfig config)
    {
        _app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
            .WithClientSecret(config.ClientSecret)
            .WithAuthority(config.Authority)
            .Build();

        // In memory token caches (App and User caches)
        _app.AddInMemoryTokenCache();
        
        _scopes = new[] { config.ApiUrl }; 
    }

    public void Dispose()
    {
        throw new NotImplementedException();
    }

    public OAuthToken GetAccessToken()
    {
        return GetAccessToken(false);
    }

    public OAuthToken GetAccessToken(bool ignoreExistingToken)
    {
        if (!ignoreExistingToken && _token != null)
        {
            return new OAuthToken(_token);
        }

        _token = GetAccessTokenAsync().GetAwaiter().GetResult();
        return new OAuthToken(_token);
    }
    
    private async Task<string?> GetAccessTokenAsync()
    {
        AuthenticationResult? result;
        
        try
        {
            result = await _app.AcquireTokenForClient(_scopes)
                .ExecuteAsync();
            
            Console.WriteLine($"The token acquired from {result.AuthenticationResultMetadata.TokenSource} {Environment.NewLine}");
        }
        catch (MsalServiceException ex)
        {
            Console.WriteLine($"Error Acquiring Token:{Environment.NewLine}{ex}{Environment.NewLine}");
            result = null;
        }

        if (result == null) return null;
        _token = result.AccessToken;
        return result.AccessToken;
    }
}