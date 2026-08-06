// Demonstrates how to strip the S/MIME signature off an Outlook message and get
// the original, readable message back.

using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.MAPI
{
    internal static class RemoveSignatureFromMsg
    {
        public static void Run()
        {
            // Sign a message first, so the example has something to unsign.
            var privateCert = new X509Certificate2(Data.Email/"MartinCertificate.pfx", "anothertestaccount");
            var msg = MapiMessage.Load(Data.Mapi/"message.msg");

            var signed = new SecureEmailManager().AttachSignature(msg, privateCert);
            Console.WriteLine($"Signed message IsSigned: {signed.IsSigned}");

            if (signed.IsSigned)
            {
                var unsignedMsg = signed.RemoveSignature();
                Console.WriteLine($"After RemoveSignature:  {unsignedMsg.IsSigned}");
                Console.WriteLine($"Subject: {unsignedMsg.Subject}");

                var outputPath = Data.Out/"RemoveSignatureFromMsg_out.msg";
                unsignedMsg.Save(outputPath);
                Console.WriteLine($"\nSaved to {outputPath}");
            }
        }
    }
}
