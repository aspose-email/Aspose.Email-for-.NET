using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Exchange_EWS
{
    class IgnoringInvalidSSLCertificates
    {
        // ExStart:IgnoringInvalidSSLCertificates
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
        // ExEnd:IgnoringInvalidSSLCertificates
    }
}
