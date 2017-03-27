using System;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.Mapi;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Exchange_EWS
{
    class SendCalendarInvitation
    {
        public static void Run()
        {
            try
            {

                // ExStart: SendCalendarInvitation
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
                // ExEnd: SendCalendarInvitation

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
