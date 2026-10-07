// Demonstrates mail merge one data row at a time.
//
// TemplateEngine.Merge(row) produces the copy for a single row, so each message can be
// sent - or checked, logged, skipped - as soon as it is built, instead of building the
// whole batch first as MailMerge does with Instantiate. That also keeps memory flat for
// large mailings.
//
// The merged messages are saved to Out. The data source holds sample addresses, so when
// SMTP is configured each copy is sent to your own address instead.

using System;
using System.Data;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Tools.Merging;

namespace Aspose.Email.Examples.SMTP
{
    internal static class RowWiseMailMerge
    {
        public static void Run()
        {
            var template = new MailMessage
            {
                From = "sender@example.com",
                Subject = "Your order #OrderId# has shipped",
                HtmlBody = "Hello #FirstName#,<br><br>Your order <strong>#OrderId#</strong> is on its way."
            };
            template.To.Add(new MailAddress("#Email#", true));

            var engine = new TemplateEngine(template);
            var outputDir = Data.OutSub("RowWiseMailMerge");

            SmtpClient client = null;
            if (ClientBuilder.IsSmtpConfigured)
                client = ClientBuilder.Smtp(AuthType.Basic);

            try
            {
                foreach (DataRow row in CreateOrders().Rows)
                {
                    var message = engine.Merge(row);
                    message.Save(outputDir/$"order-{row["OrderId"]}.eml", SaveOptions.DefaultEml);
                    Console.Write($"{message.To}: {message.Subject}");

                    if (client != null)
                    {
                        message.To.Clear();
                        message.To.Add(client.Username);
                        client.Send(message);
                        Console.Write($"  -> sent to {client.Username}");
                    }

                    Console.WriteLine();
                }
            }
            finally
            {
                client?.Dispose();
            }

            Console.WriteLine($"\nMessages saved to {outputDir}");
            if (client == null)
                Console.WriteLine("Set Smtp.HostName in clientsettings.json to also send them.");
        }

        private static DataTable CreateOrders()
        {
            var table = new DataTable();
            table.Columns.Add("Email", typeof(string));
            table.Columns.Add("FirstName", typeof(string));
            table.Columns.Add("OrderId", typeof(string));

            table.Rows.Add("alex.smith@example.com", "Alex", "A-1001");
            table.Rows.Add("john.doe@example.com", "John", "A-1002");
            table.Rows.Add("maria.garcia@example.com", "Maria", "A-1003");

            return table;
        }
    }
}
