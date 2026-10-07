// Demonstrates mail merge: one template message, one personalised copy per data row.
//
// TemplateEngine replaces #Field# placeholders with the values of the matching DataTable
// columns - in the subject, the body and the address fields - and #Routine()#
// placeholders with what a registered routine returns. Instantiate produces all copies at
// once; RowWiseMailMerge merges one row at a time instead.
//
// The merged messages are saved to Out. The data source holds sample addresses, so when
// SMTP is configured the copies are sent to your own address instead.

using System;
using System.Data;
using Aspose.Email.Tools.Merging;

namespace Aspose.Email.Examples.SMTP
{
    internal static class MailMerge
    {
        public static void Run()
        {
            var template = new MailMessage
            {
                From = "sender@example.com",
                Subject = "Hello, #FirstName#",
                HtmlBody = "Dear #FirstName# #LastName#,<br><br>" +
                           "Thank you for your interest in <strong>Aspose.Email</strong>.<br><br>" +
                           "#GetSignature()#"
            };

            // The placeholder is not a valid address yet, so skip the address check.
            template.To.Add(new MailAddress("#Email#", true));

            var engine = new TemplateEngine(template);
            engine.RegisterRoutine("GetSignature", GetSignature);

            var messages = engine.Instantiate(CreateRecipients());

            var outputDir = Data.OutSub("MailMerge");
            for (var i = 0; i < messages.Count; i++)
            {
                var outputPath = outputDir/$"message-{i + 1}.eml";
                messages[i].Save(outputPath, SaveOptions.DefaultEml);
                Console.WriteLine($"{messages[i].To}: {messages[i].Subject}");
            }

            Console.WriteLine($"\n{messages.Count} message(s) saved to {outputDir}");

            if (!ClientBuilder.IsSmtpConfigured)
            {
                Console.WriteLine("Set Smtp.HostName in clientsettings.json to also send them.");
                return;
            }

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                foreach (var message in messages)
                {
                    message.To.Clear();
                    message.To.Add(client.Username);
                }

                client.Send(messages);
                Console.WriteLine($"Sent them to {client.Username}.");
            }
        }

        private static DataTable CreateRecipients()
        {
            var table = new DataTable();
            table.Columns.Add("Email", typeof(string));
            table.Columns.Add("FirstName", typeof(string));
            table.Columns.Add("LastName", typeof(string));

            table.Rows.Add("alex.smith@example.com", "Alex", "Smith");
            table.Rows.Add("john.doe@example.com", "John", "Doe");
            table.Rows.Add("maria.garcia@example.com", "Maria", "Garcia");

            return table;
        }

        // A template routine: whatever it returns replaces #GetSignature()#.
        private static object GetSignature(object[] args)
        {
            return "Aspose.Email Team<br>" + DateTime.Now.ToShortDateString();
        }
    }
}
