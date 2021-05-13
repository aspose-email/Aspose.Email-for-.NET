using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System;
using System.Linq;
using System.Net;

namespace Aspose.Email.Live.Demos.UI
{
    public static class HttpExtensionMethods
    {
        public static bool IsPostBack(this PageModel model)
        {
            return model.Request.IsPostBack();
        }

        public static bool IsPostBack(this HttpRequest request)
        {
            var currentUrl = UriHelper.BuildAbsolute(request.Scheme, request.Host, request.PathBase, request.Path, request.QueryString);
            var referrer = request.Headers["Referer"].FirstOrDefault();

            bool isPost = string.Compare(request.Method, "POST",
               StringComparison.CurrentCultureIgnoreCase) == 0;
            if (referrer == null) return false;

            bool isSameUrl = string.Compare(currentUrl,
               referrer,
               StringComparison.CurrentCultureIgnoreCase) == 0;

            return isPost && isSameUrl;
        }

        public static T GetService<T>(this HttpContext context) =>
            (T)context.RequestServices.GetService(typeof(T));

        public static bool CheckImageExistRemotely(string uriToImage, string mimeType)
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(uriToImage);

                request.Timeout = 300;
                request.Method = "HEAD";
                var response = (HttpWebResponse)request.GetResponse();
                return response.StatusCode == HttpStatusCode.OK && response.ContentType == mimeType;
            }
            catch
            {
                return false;
            }
        }
    }
}
