// Demonstrates how to pull the body of one specific alternate view out of a
// multipart message, without walking the AlternateViews collection by hand.

using System;

namespace Aspose.Email.Examples.Email
{
    internal static class GetAlternateViewContentByMediaType
    {
        public static void Run()
        {
            var msg = MailMessage.Load(Data.Email/"Message.eml");

            Console.WriteLine($"Alternate views: {msg.AlternateViews.Count}");
            foreach (var alternateView in msg.AlternateViews)
            {
                // Since 25.12 the ContentId of an alternate view is no longer generated
                // automatically - UniqueId is the stable way to tell views apart.
                Console.WriteLine($"  {alternateView.ContentType.MediaType} (id: {alternateView.UniqueId})");
            }

            foreach (var mediaType in new[] { "text/plain", "text/html" })
            {
                var body = msg.GetAlternateViewContent(mediaType);

                Console.WriteLine($"\n--- {mediaType} ---");
                Console.WriteLine(body == null ? "(no such alternate view)" : Preview(body));
            }
        }

        private static string Preview(string body) =>
            body.Length <= 200 ? body : body.Substring(0, 200) + "...";
    }
}
