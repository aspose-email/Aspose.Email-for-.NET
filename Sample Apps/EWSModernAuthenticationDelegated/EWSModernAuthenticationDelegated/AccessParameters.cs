namespace EWSModernAuthenticationDelegated
{
    public class AccessParameters
    {
        public string TenantId { get; set; }
        public string ClientId { get; set; }
        public string RedirectUri { get; set; } = "http://localhost";
        public string[] Scopes { get; set; } = { "https://outlook.office365.com/EWS.AccessAsUser.All" };
    }
}