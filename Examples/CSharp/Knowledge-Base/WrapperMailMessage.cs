using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    // ExStart:WrapperMailMessage
    public class WrapperMailMessage
    {
        private MailMessage msg;
        public string strFileName;

        public WrapperMailMessage(string strMessageFile)
        {
            msg = null;
            strFileName = strMessageFile;
            // Now load the EML file with a timout of 4 seconds
            CallWithTimeout(LoadEmlFile, 4000);
        }

        // Function to load the provided EML file
        public void LoadEmlFile()
        {
            msg = MailMessage.Load(strFileName);
        }

        // TimeOut function that calls the load EML function with Timeout
        public void CallWithTimeout(Action action, int timeoutMilliseconds)
        {
            Thread threadToKill = null;
            Action wrappedAction = () =>
            {
                threadToKill = Thread.CurrentThread;
                action();
            };

            IAsyncResult result = wrappedAction.BeginInvoke(null, null);
            if (result.AsyncWaitHandle.WaitOne(timeoutMilliseconds))
            {
                wrappedAction.EndInvoke(result);
            }
            else
            {
                threadToKill.Abort();
            }
        }

        // This function returns the loaded EML message or NULL in case of unsuccessful load
        public MailMessage GetMailMessage()
        {
            return msg;
        }
    }
    // ExEnd:WrapperMailMessage
}
