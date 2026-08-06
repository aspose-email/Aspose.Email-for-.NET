// Demonstrates the two ways SecureEmailManager can sign a message: a detached
// signature keeps the body readable to clients that cannot verify it, a non-detached
// one wraps the whole message.

using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Email.Mapi;

namespace Aspose.Email.Examples.Email
{
    internal static class AttachSignatureToMessage
    {
        public static void Run()
        {
            var privateCert = new X509Certificate2(Data.Email/"MartinCertificate.pfx", "anothertestaccount");
            var manager = new SecureEmailManager();

            var msg = MapiMessage.Load(Data.Email/"Message.msg");
            Console.WriteLine($"Source message IsSigned: {msg.IsSigned}");

            var signedDetached = manager.AttachSignature(msg, privateCert, new SignatureOptions { Detached = true });
            Console.WriteLine($"Detached signature attached:     {signedDetached.IsSigned}");
            signedDetached.Save(Data.Out/"AttachSignatureToMessage_detached.msg");

            var signedNonDetached = manager.AttachSignature(msg, privateCert, new SignatureOptions { Detached = false });
            Console.WriteLine($"Non-detached signature attached: {signedNonDetached.IsSigned}");
            signedNonDetached.Save(Data.Out/"AttachSignatureToMessage_nondetached.msg");

            Console.WriteLine($"\nSaved both variants to {Data.Out}");
        }
    }
}
