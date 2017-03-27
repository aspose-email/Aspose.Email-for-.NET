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
    public class WrapperMailMessage_2_0
    {
        private MailMessage msg;
        public string strFileName;
        public WrapperMailMessage_2_0(string strMessageFile)
        {
            msg = null;
            strFileName = strMessageFile;
            // Now load the EML file with a timout of 4 seconds
            CallWithTimeout(new Action(LoadEmlFile), 4000);
        }

        // Function to load the provided EML file
        public void LoadEmlFile()
        {
            msg = MailMessage.Load(strFileName);
        }

        private static void CallWithTimeout(Delegate dlgt, int timeout, params object[] args)
        {
            ManualResetEvent waitHandle = new ManualResetEvent(false);
            Thread t = new Thread(new ThreadStart(delegate
            {
                dlgt.Method.Invoke(dlgt.Target, args);
                waitHandle.Set();
            }));
            t.Start();
            bool bExec = waitHandle.WaitOne(timeout);
            if (!bExec)
                t.Abort();
        }

        // This needs to be here as .NET 2.0 doesn't allow argument less calls for this
        public delegate void Action();

        // This function returns the loaded EML message or NULL in case of unsuccessful load
        public MailMessage GetMailMessage()
        {
            return msg;
        }
    }
    // ExEnd:WrapperMailMessage
}
