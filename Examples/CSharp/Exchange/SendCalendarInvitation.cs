using Aspose.Email.Exchange;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.Exchange
{
    class SendCalendarInvitation
    {
        public static void Run()
        {
            //ExStart: SendCalendarInvitation
            /// <summary>
            /// This exmpale shows how an Exchange user can share his/her calendar with someone using the EWS client of the API. 
            /// Available from Aspose.Email for .NET 6.4.0 onwards
            /// </summary>
            using (IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain"))
            {
                // delegate calendar access permission
                ExchangeDelegateUser delegateUser = new ExchangeDelegateUser("sharingfrom@domain.com", ExchangeDelegateFolderPermissionLevel.NotSpecified);
                delegateUser.FolderPermissions.CalendarFolderPermissionLevel = ExchangeDelegateFolderPermissionLevel.Reviewer;
                client.DelegateAccess(delegateUser, "sharingfrom@domain.com");

                // create invitation
                MapiMessage mapiMessage = client.CreateCalendarSharingInvitationMessage("sharingfrom@domain.com");

                // send invitation
                MailMessageInterpretor messageInterpretor = MailMessageInterpretorFactory.Instance.GetIntepretor(mapiMessage.MessageClass);
                MailMessage mailMessage = messageInterpretor.InterpretAsTnef(mapiMessage);
                client.Send(mailMessage);
            }
            //ExEnd: SendCalendarInvitation
        }
    }
}
