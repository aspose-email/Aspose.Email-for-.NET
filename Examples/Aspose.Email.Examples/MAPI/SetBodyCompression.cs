// Demonstrates how to compress the RTF body when converting a message to MSG,
// which makes the resulting file noticeably smaller.

using System;
using System.IO;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class SetBodyCompression
    {
        public static void Run()
        {
            var message = MailMessage.Load(Data.Mapi/"message.msg");

            var uncompressedPath = Save(message, useCompression: false, fileName: "SetBodyCompression_uncompressed_out.msg");
            var compressedPath = Save(message, useCompression: true, fileName: "SetBodyCompression_compressed_out.msg");

            var uncompressed = new FileInfo(uncompressedPath).Length;
            var compressed = new FileInfo(compressedPath).Length;

            Console.WriteLine($"Without compression: {uncompressed,8:N0} bytes  {uncompressedPath}");
            Console.WriteLine($"With compression:    {compressed,8:N0} bytes  {compressedPath}");
            Console.WriteLine($"Saved {uncompressed - compressed:N0} bytes.");
        }

        private static string Save(MailMessage message, bool useCompression, string fileName)
        {
            var options = new MapiConversionOptions { UseBodyCompression = useCompression };
            var mapi = MapiMessage.FromMailMessage(message, options);

            var path = Data.Out/fileName;
            mapi.Save(path);
            return path;
        }
    }
}
