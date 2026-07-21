using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class GetFolderTypeInformationUsingEWS
    {
        public static void Run()
        {
            const string mailboxUri = "https://exchange/ews/exchange.asmx";
            const string domain = @"";
            const string username = @"username@ASE305.onmicrosoft.com";
            const string password = @"password";
            NetworkCredential credentials = new NetworkCredential(username, password, domain);

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
            }
        }
    }
}
