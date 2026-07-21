using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace Aspose.Email.Examples.EWS
{
    class IgnoringInvalidSSLCertificates
    {
        public static void Run()
        {
            // Register callback method for SSL validation event
            ServicePointManager.ServerCertificateValidationCallback += RemoteCertificateValidationHandler;
        }

        // This event handler is called when SSL certificate is verified
        private static bool RemoteCertificateValidationHandler(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true; //Ignore the checks and go ahead
        }
    }
}
