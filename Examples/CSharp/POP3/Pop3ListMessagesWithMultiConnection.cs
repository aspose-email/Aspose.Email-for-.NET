using System;
using Aspose.Email.Clients;
using Aspose.Email.Clients.Pop3;

namespace Aspose.Email.Examples.CSharp.Email.POP3
{
    class Pop3ListMessagesWithMultiConnection
    {
        public static void Run()
        {
            //ExStart:1
            // Create an instance of the Pop3Client class
            Pop3Client pop3Client = new Pop3Client();
            pop3Client.Host = "<HOST>";
            pop3Client.Port = 995;
            pop3Client.Username = "<USERNAME>";
            pop3Client.Password = "<PASSWORD>";

            pop3Client.ConnectionsQuantity = 5;
            pop3Client.UseMultiConnection = MultiConnectionMode.Enable;
            DateTime multiConnectionModeStartTime = DateTime.Now;
            Pop3MessageInfoCollection messageInfoCol1 = pop3Client.ListMessages();
            TimeSpan multiConnectionModeTimeSpan = DateTime.Now - multiConnectionModeStartTime;

            pop3Client.UseMultiConnection = MultiConnectionMode.Disable;
            DateTime singleConnectionModeStartTime = DateTime.Now;
            Pop3MessageInfoCollection messageInfoCol2 = pop3Client.ListMessages();
            TimeSpan singleConnectionModeTimeSpan = DateTime.Now - singleConnectionModeStartTime;

            Console.WriteLine("multyConnectionModeTimeSpan: " + multiConnectionModeTimeSpan.TotalMilliseconds);
            Console.WriteLine("singleConnectionModeTimeSpan: " + singleConnectionModeTimeSpan.TotalMilliseconds);
            double performanceRelation = singleConnectionModeTimeSpan.TotalMilliseconds / multiConnectionModeTimeSpan.TotalMilliseconds;
            Console.WriteLine("Performance Relation: " + performanceRelation);
            //ExEnd: 1

            Console.WriteLine("Pop3ListMessagesWithMultiConnection executed successfully.");
        }
    }
}
