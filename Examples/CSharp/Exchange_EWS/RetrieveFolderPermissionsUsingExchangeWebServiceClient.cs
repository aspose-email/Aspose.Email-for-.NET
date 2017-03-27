using System;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mime;
using Aspose.Email.Clients.Exchange;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class RetrieveFolderPermissionsUsingExchangeWebServiceClient
    {
        public static void Run()
        {
            // ExStart:RetrieveFolderPermissionsUsingExchangeWebServiceClient
            string folderName = "DesiredFolderName";

            // Create instance of EWSClient class by giving credentials
            IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

            ExchangeFolderInfoCollection folders = client.ListPublicFolders();
            ExchangeFolderPermissionCollection permissions = new ExchangeFolderPermissionCollection();

            ExchangeFolderInfo publicFolder = null;
            try
            {
                foreach (ExchangeFolderInfo folderInfo in folders)
                    if (folderInfo.DisplayName.Equals(folderName))
                        publicFolder = folderInfo;

                if (publicFolder == null)
                    Console.WriteLine("public folder was not created in the root public folder");

                ExchangePermissionCollection folderPermissionCol = client.GetFolderPermissions(publicFolder.Uri);
                foreach (ExchangeBasePermission perm in folderPermissionCol)
                {
                    ExchangeFolderPermission permission = perm as ExchangeFolderPermission;
                    if (permission == null)
                        Console.WriteLine("Permission is null.");
                    else
                    {
                        Console.WriteLine("User's primary smtp address: {0}", permission.UserInfo.PrimarySmtpAddress);
                        Console.WriteLine("User can create Items: {0}", permission.CanCreateItems.ToString());
                        Console.WriteLine("User can delete Items: {0}", permission.DeleteItems.ToString());
                        Console.WriteLine("Is Folder Visible: {0}", permission.IsFolderVisible.ToString());
                        Console.WriteLine("Is User owner of this folder: {0}", permission.IsFolderOwner.ToString());
                        Console.WriteLine("User can read items: {0}", permission.ReadItems.ToString());
                    }
                }
                ExchangeMailboxInfo mailboxInfo = client.GetMailboxInfo();
                //Get the Permissions for the Contacts and Calendar Folder
                ExchangePermissionCollection contactsPermissionCol = client.GetFolderPermissions(mailboxInfo.ContactsUri);
                ExchangePermissionCollection calendarPermissionCol = client.GetFolderPermissions(mailboxInfo.CalendarUri);
            }
            finally
            {
                //Do the needfull
            }
            // ExEnd:RetrieveFolderPermissionsUsingExchangeWebServiceClient
        }
    }
}