// Demonstrates tracking where a sent message went (Exchange message tracking).
//
// FindMessageTrackingReport searches the tracking logs for messages by sender,
// recipient, subject or time; each hit carries a report id. GetMessageTrackingReport
// turns that id into the delivery events per recipient - submitted, delivered, failed
// and on which server. Message tracking over EWS is an on-premises Exchange feature
// (2010 and later); the scope "Organization" searches the whole organisation.
// Replace the addresses and domain below with ones of your organisation.

using System;
using Aspose.Email.Clients.Exchange.WebService;

namespace Aspose.Email.Examples.EWS
{
    internal static class TrackMessageDelivery
    {
        public static void Run()
        {
            const string scope = "Organization";
            const string domain = "example.com";

            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                var findOptions = new FindMessageTrackingReportOptions(scope, domain)
                {
                    Sender = new MailAddress("sender@example.com"),
                    StartDateTime = DateTime.Now.AddDays(-2),
                    EndDateTime = DateTime.Now
                };

                var found = client.FindMessageTrackingReport(findOptions);
                Console.WriteLine($"{found.Length} message(s) sent by {findOptions.Sender} in the last two days.");

                foreach (var info in found)
                {
                    Console.WriteLine($"\n{info.SubmittedTime:g}  {info.Subject}");

                    var reportOptions = new GetMessageTrackingReportOptions(scope, info.MessageTrackingReportId)
                    {
                        ReportTemplate = MessageTrackingReportTemplate.Summary
                    };

                    var report = client.GetMessageTrackingReport(reportOptions);
                    foreach (var trackingEvent in report.RecipientTrackingEvents)
                    {
                        Console.WriteLine($"  {trackingEvent.Recipient}: {trackingEvent.DeliveryStatus} " +
                                          $"({trackingEvent.EventDescription}, {trackingEvent.Date:g})");
                    }
                }
            }
        }
    }
}
