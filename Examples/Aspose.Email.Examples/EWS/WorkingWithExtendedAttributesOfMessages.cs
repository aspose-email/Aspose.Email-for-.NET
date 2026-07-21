using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class WorkingWithExtendedAttributesOfMessages
    {
        public static void Run()
        {
            // Set Exchange Server web service URL, Username, password, domain information
            string mailboxURI = "https://ex2010/ews/exchange.asmx";
            string username = "test.exchange";
            string password = "pwd";
            string domain = "ex2010.local";

            // Connect to the Exchange Server
            NetworkCredential credential = new NetworkCredential(username, password, domain);
            IEWSClient client = EWSClient.GetEWSClient(mailboxURI, credential);
            {
                try
                {
                    //Create a new Property
                    PidNamePropertyDescriptor pd = new PidNamePropertyDescriptor(
                        "MyTestProp",
                        PropertyDataType.String,
                        KnownPropertySets.PublicStrings);
                    string value = "MyTestPropValue";

                    //Create a message
                    MapiMessage message = new MapiMessage(
                        "from@domain.com",
                        "to@domain.com",
                        "EMAILNET-38844 - " + Guid.NewGuid().ToString(),
                        "EMAILNET-38844 EWS: Support for create, retrieve and update Extended Attributes for Emails");

                    //Set property on the message
                    message.SetProperty(pd, value);

                    //append the message to server
                    string uri = client.AppendMessage(message);

                    //Fetch the message from server
                    MapiMessage mapiMessage = client.FetchItem(uri, new PropertyDescriptor[] { pd });

                    //Retreive the value from Message
                    string fetchedValue = mapiMessage.NamedProperties[pd].GetString();
                }
                finally
                {
                }
            }
        }
    }
}
