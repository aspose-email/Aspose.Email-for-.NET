using Aspose.Email.Clients.Exchange;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace CSharp.Exchange_EWS
{
    public class AutoDiscoverUsingEWS
    {
        public static void Run()
        {
            //ExStart: AutoDiscoverUsingEWS
            string email = "asposeemail.test3@aspose.com";
            string password = "Aspose@2017";
            AutodiscoverService svc = new AutodiscoverService();
            svc.Credentials = new NetworkCredential(email, password);

            IDictionary<UserSettingName, object> userSettings = svc.GetUserSettings(email, UserSettingName.ExternalEwsUrl).Settings;
            string ewsUrl = (string)userSettings[UserSettingName.ExternalEwsUrl];
            Console.WriteLine("Auto discovered EWS Url: "  + ewsUrl);
            //ExEnd: AutoDiscoverUsingEWS
        }
    }
}
