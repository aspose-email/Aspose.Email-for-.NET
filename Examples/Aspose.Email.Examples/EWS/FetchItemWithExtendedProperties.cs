using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Clients.Exchange.WebService.Models;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Email.Examples.EWS
{
    internal class FetchItemWithExtendedProperties
    {
        public static void Run()
        {
            // Create instance of ExchangeWebServiceClient class by giving credentials
            var ewsClient = ClientBuilder.Ews(AuthType.ModernWithAppPermission);

            //// Define an extended property
            //var extendedProperty = KnownPropertyList.BodyContentId;

            //// Create a message and set an extended property value
            //var msg = new MapiMessage("from@from.com", "to@to.com", "Test message", "This is a test message");
            //msg.SetProperty(extendedProperty, "Fghtwujdb.dkeheyo.ghgwo");

            //// Append message
            //var uri = ewsClient.AppendMessage(ewsClient.MailboxInfo.InboxUri, msg, true);

            //// Fetch appended item. Pass the extended property descriptor as method parameter.
            //var fetchedMsg = ewsClient.FetchItem(uri, new PropertyDescriptor[] { extendedProperty });

            //// Print the extended property value
            //Console.WriteLine(fetchedMsg.Properties[extendedProperty].GetString());

            //var itemUris = ewsClient.ListItems(ewsClient.MailboxInfo.InboxUri);

            

            var messageInfoList = ewsClient.ListMessages(ewsClient.MailboxInfo.InboxUri);
            var uriList = messageInfoList.Select(item => item.UniqueUri).ToList();
            var options = EwsFetchItems.Create().AddUris(uriList);
            var items = ewsClient.FetchItems(options);
        }
    }
}
