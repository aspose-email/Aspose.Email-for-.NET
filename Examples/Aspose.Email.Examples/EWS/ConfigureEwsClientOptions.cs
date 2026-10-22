// Demonstrates the transport settings of the EWS client.
//
// - ReconnectCount: how often the client retries when the connection breaks.
// - EnableDecompression: accept compressed responses, which shrinks large listings on
//   slow links.
// - ReturnClientRequestId: ask Exchange to return the request id, which helps to match
//   a request with the server's logs when troubleshooting.
// - AddHeader / Headers / RemoveHeader: custom HTTP headers sent with every request,
//   for example to route requests through a proxy or tag them in logs.

using System;

namespace Aspose.Email.Examples.EWS
{
    internal static class ConfigureEwsClientOptions
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Ews(AuthType.ModernWithAppPermission))
            {
                client.ReconnectCount = 3;
                client.EnableDecompression = true;
                client.ReturnClientRequestId = true;

                client.AddHeader("X-Client-Application", "Aspose.Email Examples");
                client.AddHeader("X-Temporary-Tag", "example");
                client.RemoveHeader("X-Temporary-Tag");

                Console.WriteLine($"Reconnect attempts:      {client.ReconnectCount}");
                Console.WriteLine($"Decompression enabled:   {client.EnableDecompression}");
                Console.WriteLine($"Return client request id: {client.ReturnClientRequestId}");
                Console.WriteLine("Custom headers:");
                foreach (var header in client.Headers)
                    Console.WriteLine($"  {header.Key}: {header.Value}");

                var messages = client.ListMessages(client.MailboxInfo.InboxUri, 5);
                Console.WriteLine($"\nListed {messages.Count} message(s) with these settings.");
            }
        }
    }
}
