using System;
using System.Threading;
using Aspose.Email.Exchange;
using Aspose.Email.Mail;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange
{
    class GetContactsExtendedProperties
    {
        public static void Run()
        {
            try
            {
                // ExStart:GetContactsExtendedProperties
                IEWSClient client = EWSClient.GetEWSClient("https://outlook.office365.com/ews/exchange.asmx", "testUser", "pwd", "domain");

                // The required extended properties must be added in order to create or read them from the Exchange Server
                string[] extraFields = new string[] { "User Field 1", "User Field 2", "User Field 3", "User Field 4" };
                foreach (string extraField in extraFields)
                    client.ContactExtendedPropertiesDefinition.Add(extraField);

                // Create a test contact on the Exchange Server
                Contact contact = new Contact();
                contact.DisplayName = "EMAILNET-38433 - " + Guid.NewGuid().ToString();
                foreach (string extraField in extraFields)
                    contact.ExtendedProperties.Add(extraField, extraField);

                string contactId = client.CreateContact(contact);

                // retrieve the contact back from the server after some time
                Thread.Sleep(5000);
                contact = client.GetContact(contactId);

                // Parse the extended properties of contact 
                foreach (string extraField in extraFields)
                    if (contact.ExtendedProperties.ContainsKey(extraField))
                        Console.WriteLine(contact.ExtendedProperties[extraField].ToString());
                // ExEnd:GetContactsExtendedProperties
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
                throw;
            }
        }
    }
}