using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Mapi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class RetreiveExtAttributesForCalendarItems
    {
        public static void Run()
        {
            IEWSClient client = EWSClient.GetEWSClient("https://exchange.office365.com/Exchange.asmx", "username","password");
            
            //Fetch all calendars from Exchange calendar's folder
            string[] uriList = client.ListItems(client.MailboxInfo.CalendarUri);

            //Define the Extended Attribute Property Descriptor for searching purpose
            //In this case, we have a K1 Long named property for Calendar item
            PropertyDescriptor propertyDescriptor = new PidNamePropertyDescriptor("K1", PropertyDataType.Integer32, new Guid("00020329-0000-0000-C000-000000000046"));

            //Fetch calendars that have the custom property
            IList<MapiCalendar> mapiCalendarList = client.FetchMapiCalendar(uriList, new PropertyDescriptor[] { propertyDescriptor });

            foreach (MapiCalendar cal in mapiCalendarList)
            {
                foreach (MapiNamedProperty namedProperty in cal.NamedProperties.Values)
                {
                    Console.WriteLine(namedProperty.NameId + " = " + namedProperty.GetInt32());
                }
            }
        }
    }
}
