using Aspose.Email.Calendar;
using Aspose.Email.Clients.Exchange;
using Aspose.Email.Clients.Exchange.WebService;
using Aspose.Email.Tools.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class FilterAppointmentsUsingEWS
    {
        public static void Run()
        {
            // Connect to EWS
            const string mailboxUri = "https://outlook.office365.com/ews/exchange.asmx";
            const string username = "username";
            const string password = "password";
            const string domain = "domain";

            IEWSClient client = EWSClient.GetEWSClient(mailboxUri, username, password, domain);

            DateTime startTime = new DateTime(2017,09, 15);
            DateTime endTime = new DateTime(2017, 10, 10);
            ExchangeQueryBuilder builder = new ExchangeQueryBuilder();
            builder.Appointment.Start.Since(startTime);
            builder.Appointment.End.BeforeOrEqual(endTime);
            MailQuery query = builder.GetQuery();
            Appointment[] appointments = client.ListAppointments(query);

            builder = new ExchangeQueryBuilder();
            builder.Appointment.IsRecurring.Equals(false);
            query = builder.GetQuery();
            appointments = client.ListAppointments(query);
        }
    }
}
