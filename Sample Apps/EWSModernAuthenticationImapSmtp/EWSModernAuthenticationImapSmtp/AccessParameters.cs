namespace EWSModernAuthenticationImapSmtp
{
    public class AccessParameters
    {
        public string UserName { get; set; }
        public string TenantId { get; set; }
        public string ClientId { get; set; }
        public string RedirectUri { get; set; } = "http://localhost";
        public string[] Scopes { get; set; } = { "https://outlook.office.com/IMAP.AccessAsUser.All", "https://outlook.office.com/SMTP.Send" };
    }
}