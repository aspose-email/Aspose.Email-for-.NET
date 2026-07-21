using System;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.EWS
{
    class SendCalendarInvitation
    {
        public static void Run()
        {
            try
            {

                using (IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain"))
                {
                    // delegate calendar access permission
                    ExchangeDelegateUser delegateUser = new ExchangeDelegateUser("sharingfrom@domain.com", ExchangeDelegateFolderPermissionLevel.NotSpecified);
                    delegateUser.FolderPermissions.CalendarFolderPermissionLevel = ExchangeDelegateFolderPermissionLevel.Reviewer;
                    client.DelegateAccess(delegateUser, "sharingfrom@domain.com");

                    // Create invitation
                    MapiMessage mapiMessage = client.CreateCalendarSharingInvitationMessage("sharingfrom@domain.com");
                    MailConversionOptions options = new MailConversionOptions();
                    options.ConvertAsTnef = true;
                    MailMessage mail = mapiMessage.ToMailMessage(options);
                    client.Send(mail);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
