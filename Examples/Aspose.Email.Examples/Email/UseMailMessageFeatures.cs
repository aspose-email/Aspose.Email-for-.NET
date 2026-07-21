// Demonstrates how to set message priority, sensitivity, date, and delivery notification
// options on a MailMessage and save it as EML.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class UseMailMessageFeatures
    {
        public static void Run()
        {
            var eml = new MailMessage
            {
                From = "sender@gmail.com",
                To = "receiver@gmail.com",
                Subject = "Using MailMessage Features",
                Date = DateTime.Now,
                Priority = MailPriority.High,
                Sensitivity = MailSensitivity.Normal,
                DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess
            };

            var outputPath = Data.Out/"UseMailMessageFeatures_out.eml";
            eml.Save(outputPath, SaveOptions.DefaultEml);

            // Each of these becomes a header the receiving client acts on.
            Console.WriteLine($"Priority:     {eml.Priority}");
            Console.WriteLine($"Sensitivity:  {eml.Sensitivity}");
            Console.WriteLine($"Notification: {eml.DeliveryNotificationOptions}");
            Console.WriteLine($"Saved to {outputPath}");
        }
    }
}
