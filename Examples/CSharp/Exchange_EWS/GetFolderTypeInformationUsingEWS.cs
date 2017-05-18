using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class GetFolderTypeInformationUsingEWS
    {
        public static void Run()
        {
            //ExStart: GetFolderTypeInformationUsingEWS
            const string mailboxUri = "https://exchange/ews/exchange.asmx";
            const string domain = @"";
            const string username = @"username@ASE305.onmicrosoft.com";
            const string password = @"password";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);

            // ExStart:CopyConversations
            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, credentials);

            ExchangeFolderInfoCollection folderInfoCol = client.ListSubFolders(client.MailboxInfo.RootUri);
            foreach (ExchangeFolderInfo folderInfo in folderInfoCol)
            {
                switch (folderInfo.FolderType)
                {
                    case ExchangeFolderType.Appointment:
                        // handle Appointment
                        break;
                    case ExchangeFolderType.Contact:
                        // handle Contact
                        break;
                    case ExchangeFolderType.Task:
                        // handle Task
                        break;
                    case ExchangeFolderType.Note:
                        // handle email message
                        break;
                    case ExchangeFolderType.StickyNote:
                        // handle StickyNote
                        break;
                    case ExchangeFolderType.Journal:
                        // handle Journal
                        break;
                    default:
                        break;
                }
                //ExEnd: GetFolderTypeInformationUsingEWS
            }
        }
    }
}
