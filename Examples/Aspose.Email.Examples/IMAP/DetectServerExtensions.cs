// Demonstrates how to find out which optional IMAP extensions the server supports.
//
// Much of the ImapClient API maps onto extensions a server may or may not implement -
// MOVE, CONDSTORE, SORT, THREAD, QUOTA and so on. The client reads the server's
// CAPABILITY list and exposes it as a set of *Supported properties, so you can check
// before calling a method that depends on one. RetreivingServerExtensions prints the
// raw list instead.

using System;
using Aspose.Email.Clients.Imap;

namespace Aspose.Email.Examples.IMAP
{
    internal static class DetectServerExtensions
    {
        public static void Run()
        {
            using (var client = ClientBuilder.Imap(AuthType.ModernWithDelegatedPermission))
            {
                // Connects the client and loads the capability list.
                var capabilities = client.GetCapabilities();
                Console.WriteLine($"{client.Host} announces {capabilities.Length} capabilities.\n");

                Show("ID              (RFC 2971)", client.IdSupported);
                Show("NAMESPACE       (RFC 2342)", client.NamespaceSupported);
                Show("QUOTA           (RFC 2087)", client.QuotaSupported);
                Show("UIDPLUS         (RFC 4315)", client.UidPlusSupported);
                Show("SASL-IR         (RFC 4959)", client.SaslIrSupported);
                Show("ENABLE          (RFC 5161)", client.EnableSupported);
                Show("UNSELECT        (RFC 3691)", client.UnselectSupported);
                Show("CHILDREN        (RFC 3348)", client.ChildrenSupported);
                Show("LIST-EXTENDED   (RFC 5258)", client.ExtendedListSupported);
                Show("SPECIAL-USE     (RFC 6154)", client.SpecialUseSupported);
                Show("MOVE            (RFC 6851)", client.MoveSupported);
                Show("CONDSTORE       (RFC 7162)", client.CondstoreSupported);
                Show("QRESYNC         (RFC 7162)", client.QresyncSupported);
                Show("ESEARCH         (RFC 4731)", client.ESearchSupported);
                Show("ANNOTATE        (RFC 5257)", client.AnnotateSupported);
                Show("SORT            (RFC 5256)", client.SortSupported);
                Show("THREAD          (RFC 5256)", client.ThreadSupported);
                Show("COMPRESS        (RFC 4978)", client.CompressSupported);
                Show("X-GM-EXT-1      (Gmail)", client.GmExt1Supported);

                if (client.ThreadSupported)
                    Console.WriteLine($"\nThread algorithms: {string.Join(", ", client.ThreadAlgorithms)}");

                if (client.CompressSupported)
                    Console.WriteLine($"Compression:       {client.ServerSupportedCompression}");
            }
        }

        private static void Show(string extension, bool supported)
        {
            Console.WriteLine($"  {extension,-28} {(supported ? "yes" : "-")}");
        }
    }
}
