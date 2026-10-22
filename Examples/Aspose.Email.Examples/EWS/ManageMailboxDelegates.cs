// Demonstrates managing delegates - people allowed to work in someone else's mailbox,
// such as an assistant who handles a manager's calendar.
//
// ListDelegates shows who has access and with which permission level per folder.
// DelegateAccess grants access, UpdateDelegate changes it, and CloseAccess revokes it.
// Replace the addresses below with a mailbox and a user of your organisation.

using System;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class ManageMailboxDelegates
    {
        public static void Run()
        {
            const string mailbox = "manager@example.com";
            const string assistant = "assistant@example.com";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                PrintDelegates(client, mailbox, "Before:");

                // Grant: the assistant may read and edit the calendar and gets meeting messages.
                var delegateUser = new ExchangeDelegateUser
                {
                    UserInfo = new ExchangeFolderUserInfo { PrimarySmtpAddress = assistant },
                    FolderPermissions = new ExchangeDelegatePermissions
                    {
                        CalendarFolderPermissionLevel = ExchangeDelegateFolderPermissionLevel.Editor
                    },
                    Options = ExchangeDelegateOptions.ReceiveCopiesOfMeetingMessages
                };
                client.DelegateAccess(delegateUser, mailbox);
                PrintDelegates(client, mailbox, "After granting access:");

                // Update: add read access to the Inbox.
                delegateUser.FolderPermissions.InboxFolderPermissionLevel =
                    ExchangeDelegateFolderPermissionLevel.Reviewer;
                client.UpdateDelegate(delegateUser, mailbox);
                PrintDelegates(client, mailbox, "After updating access:");

                // Revoke.
                client.CloseAccess(assistant, mailbox);
                PrintDelegates(client, mailbox, "After revoking access:");
            }
        }

        private static void PrintDelegates(IEWSClient client, string mailbox, string title)
        {
            var delegates = client.ListDelegates(mailbox);

            Console.WriteLine($"{title} {delegates.Count} delegate(s) of {mailbox}");
            foreach (var user in delegates)
            {
                var permissions = user.FolderPermissions;
                Console.WriteLine($"  {user.UserInfo?.PrimarySmtpAddress}: " +
                                  $"calendar {permissions?.CalendarFolderPermissionLevel}, " +
                                  $"inbox {permissions?.InboxFolderPermissionLevel}, options {user.Options}");
            }
        }
    }
}
