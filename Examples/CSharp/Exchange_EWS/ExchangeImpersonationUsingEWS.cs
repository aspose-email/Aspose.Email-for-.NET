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
    class ExchangeImpersonationUsingEWS
    {
        public static void Run()
        {
            // ExStart:ExchangeImpersonationUsingEWS
            // Create instance's of EWSClient class by giving credentials
            IEWSClient client1 = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser1", "pwd", "domain");
            IEWSClient client2 = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser2", "pwd", "domain");
            {
                string folder = "Drafts";
                try
                {
                    foreach (ExchangeMessageInfo messageInfo in client1.ListMessages(folder))
                        client1.DeleteMessage(messageInfo.UniqueUri);
                    string subj1 = string.Format("NETWORKNET_33354 {0} {1}", "User", "User1");
                    client1.AppendMessage(folder, new MailMessage("User1@exchange.conholdate.local", "To@aspsoe.com", subj1, ""));

                    foreach (ExchangeMessageInfo messageInfo in client2.ListMessages(folder))
                        client2.DeleteMessage(messageInfo.UniqueUri);
                    string subj2 = string.Format("NETWORKNET_33354 {0} {1}", "User", "User2");
                    client2.AppendMessage(folder, new MailMessage("User2@exchange.conholdate.local", "To@aspose.com", subj2, ""));

                    ExchangeMessageInfoCollection messInfoColl = client1.ListMessages(folder);
                    //Assert.AreEqual(1, messInfoColl.Count);
                    //Assert.AreEqual(subj1, messInfoColl[0].Subject);

                    client1.ImpersonateUser(ItemChoice.PrimarySmtpAddress, "User2@exchange.conholdate.local");
                    ExchangeMessageInfoCollection messInfoColl1 = client1.ListMessages(folder);
                    //Assert.AreEqual(1, messInfoColl1.Count);
                    //Assert.AreEqual(subj2, messInfoColl1[0].Subject);

                    client1.ResetImpersonation();
                    ExchangeMessageInfoCollection messInfoColl2 = client1.ListMessages(folder);
                    //Assert.AreEqual(1, messInfoColl2.Count);
                    //Assert.AreEqual(subj1, messInfoColl2[0].Subject);
                }
                finally
                {
                    try
                    {
                        foreach (ExchangeMessageInfo messageInfo in client1.ListMessages(folder))
                            client1.DeleteMessage(messageInfo.UniqueUri);
                        foreach (ExchangeMessageInfo messageInfo in client2.ListMessages(folder))
                            client2.DeleteMessage(messageInfo.UniqueUri);
                    }
                    catch { }
                }
            }
            // ExEnd:ExchangeImpersonationUsingEWS
        }
    }
}