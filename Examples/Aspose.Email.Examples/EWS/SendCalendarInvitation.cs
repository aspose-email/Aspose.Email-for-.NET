// Demonstrates sharing your calendar with a colleague.
//
// Sharing has two parts: DelegateAccess grants the colleague permission on the calendar
// folder (here Reviewer, read-only), and CreateCalendarSharingInvitationMessage builds
// the invitation Outlook shows with an "Open this calendar" button. It must be sent as
// TNEF to keep the sharing data. Replace the addresses below with people of your
// organisation; the example revokes the access again at the end.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.EWS
{
    internal static class SendCalendarInvitation
    {
        public static void Run()
        {
            const string owner = "owner@example.com";
            const string colleague = "colleague@example.com";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var delegateUser = new ExchangeDelegateUser
                {
                    UserInfo = new ExchangeFolderUserInfo { PrimarySmtpAddress = colleague },
                    FolderPermissions = new ExchangeDelegatePermissions
                    {
                        CalendarFolderPermissionLevel = ExchangeDelegateFolderPermissionLevel.Reviewer
                    }
                };
                client.DelegateAccess(delegateUser, owner);

                try
                {
                    var invitation = client.CreateCalendarSharingInvitationMessage(colleague);
                    var message = invitation.ToMailMessage(new MailConversionOptions { ConvertAsTnef = true });

                    client.Send(message);
                    Console.WriteLine($"Shared the calendar of {owner} with {colleague} and sent the invitation.");
                }
                finally
                {
                    client.CloseAccess(colleague, owner);
                    Console.WriteLine("Access revoked again.");
                }
            }
        }
    }
}
