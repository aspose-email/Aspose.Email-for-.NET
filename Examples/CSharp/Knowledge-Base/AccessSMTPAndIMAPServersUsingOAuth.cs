using Aspose.Email.Clients;
using Aspose.Email.Clients.Imap;
using Aspose.Email.Clients.Smtp;
using Aspose.Email.Examples.CSharp.Email.Gmail;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    class AccessSMTPAndIMAPServersUsingOAuth
    {
        // ExStart:AccessSMTPAndIMAPServersUsingOAuth
        public static void Run()
        {
            try
            {
                // The authorizationCode should be replaced with your value. To get authorizationCode use the URL bellow: https://accounts.google.com/o/oauth2/auth?redirect_uri=urn:ietf:wg:oauth:2.0:oob&response_type=code&client_id=929645059575.apps.googleusercontent.com&scope=https%3A%2F%2Fmail.google.com%2F
                string authorizationCode = "4/zx4_I4ZhhUdmgLdsejpjeMkwAAMs.kk7o1Qx9U28VOl05ti8ZT3Y1uIlidQI"; // authorizationCode should be replaced with new value !!!!!!!
                string accessToken = GetAccessToken(authorizationCode);
                AccessSMTPServer(accessToken);
                AccessIMAPServer(accessToken);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static void AccessSMTPServer(string accessToken)
        {
            MailMessage message = new MailMessage("user1@testaccount1913.narod2.ru", "user1@testaccount1913.narod2.ru", "NETWORKNET-33499 - " + Guid.NewGuid().ToString(), "Access to SMTP servers using OAuth");
            using (SmtpClient client = new SmtpClient("smtp.gmail.com", 587, "user1@testaccount1913.narod2.ru", accessToken, true))
            {
                client.Timeout = 400000;
                client.SecurityOptions = SecurityOptions.SSLImplicit;
                client.Send(message);
            }
        }

        static void AccessIMAPServer(string accessToken)
        {
            using (ImapClient client = new ImapClient("imap.gmail.com", 993, "user1@testaccount1913.narod2.ru", accessToken, true))
            {
                client.SecurityOptions= SecurityOptions.SSLImplicit;                
                client.SelectFolder("Inbox");
                ImapMessageInfoCollection messageInfoCol = client.ListMessages();
            }
        }

        internal static string GetAccessToken(string authorizationCode)
        {
            string actionUrl = "https://accounts.google.com/o/oauth2/token";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(actionUrl);
            request.CookieContainer = new CookieContainer();
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            string encodedParameters = string.Format("client_id={1}&code={0}&client_secret={2}&redirect_uri={3}&grant_type={4}",
            HttpUtility.UrlEncode(authorizationCode),
            HttpUtility.UrlEncode("929645059575.apps.googleusercontent.com"),
            HttpUtility.UrlEncode("USnH5eQRsC4XrjJbpGG7WVq5"),
            HttpUtility.UrlEncode("urn:ietf:wg:oauth:2.0:oob"),
            HttpUtility.UrlEncode("authorization_code"));
            byte[] requestData = Encoding.UTF8.GetBytes(encodedParameters);
            request.ContentLength = requestData.Length;
            if (requestData.Length > 0)
                using (Stream stream = request.GetRequestStream())
                    stream.Write(requestData, 0, requestData.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            string responseText = null;
            using (TextReader reader = new StreamReader(response.GetResponseStream(), Encoding.ASCII))
                responseText = reader.ReadToEnd();
            string accessToken = null;
            foreach (string sPair in responseText.Replace("{", "").Replace("}", "").Replace("\"", "").
            Split(new string[] { ",\n" }, StringSplitOptions.None))
            {
                string[] pair = sPair.Split(':');
                if ("access_token" == pair[0].Trim())
                {
                    accessToken = pair[1].Trim();
                    break;
                }
            }
            return accessToken;
        }
        // ExEnd:AccessSMTPAndIMAPServersUsingOAuth
    }
}
