namespace GraphApp;
using Microsoft.Extensions.Configuration;
using System.Globalization;

public class AuthenticationConfig
{
    public string Instance { get; set; }
        
    public string ApiUrl { get; set; }
        
    public string TenantId { get; set; }
        
    public string ClientId { get; set; }
        
    public string UserId { get; set; }
        
    public string Authority => string.Format(CultureInfo.InvariantCulture, Instance, TenantId);

    public string ClientSecret { get; set; }
        
    public static AuthenticationConfig ReadFromJsonFile(string path)
    {
        var builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile(path);

        var configuration = builder.Build();
        return configuration.Get<AuthenticationConfig>();
    }
}