namespace EWSModernAuthenticationApp
{
    public class AccessParameters
    {
        public string UserName { get; set; }
        public string TenantId { get; set; }
        public string AppId { get; set; }
        public string ClientSecret { get; set; }
        public string[] Scopes { get; set; } = { "https://outlook.office365.com/.default" };
    }
}