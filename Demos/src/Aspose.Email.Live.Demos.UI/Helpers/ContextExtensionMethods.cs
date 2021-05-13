using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Extensions;
using Microsoft.Extensions.Caching.Memory;
using System;
using System.Linq;

namespace Aspose.Email.Live.Demos.UI
{
    public static class ContextExtensionMethods
    {
        public static IMemoryCache GetCache(this HttpContext context)
        {
            return context.GetService<IMemoryCache>();
        }

        public static Uri GetUrl(this HttpRequest request)
        {
            return new Uri(UriHelper.GetDisplayUrl(request));
        }

        public static string[] GetUserLanguages(this HttpRequest request)
        {
            return request.GetTypedHeaders()
                .AcceptLanguage
                ?.OrderByDescending(x => x.Quality ?? 1)
                .Select(x => x.Value.ToString())
                .ToArray() ?? Array.Empty<string>();
        }
    }
}
