using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Primitives;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace Aspose.Email.Live.Demos.UI.Helpers
{
    public class HttpResponseMessageResult : IActionResult
    {
        public HttpResponseMessage ResponseMessage { get; }

        public HttpResponseMessageResult(HttpResponseMessage responseMessage)
        {
            ResponseMessage = responseMessage; // could add throw if null
        }

        public HttpResponseMessageResult(HttpStatusCode statusCode)
        {
            ResponseMessage = new HttpResponseMessage(statusCode);
        }

        public async Task ExecuteResultAsync(ActionContext context)
        {
            context.HttpContext.Response.StatusCode = (int)ResponseMessage.StatusCode;

            foreach (var header in ResponseMessage.Headers)
            {
                context.HttpContext.Response.Headers.TryAdd(header.Key, new StringValues(header.Value.ToArray()));
            }

            if (ResponseMessage.Content == null)
                return;

            using (var stream = await ResponseMessage.Content.ReadAsStreamAsync())
            {
                context.HttpContext.Response.ContentLength = ResponseMessage.Content.Headers.ContentLength;

                if (context.HttpContext.Response.ContentLength > 0)
                    context.HttpContext.Response.ContentType = ResponseMessage.Content.Headers.ContentType.ToString();

                foreach (var header in ResponseMessage.Content.Headers)
                {
                    context.HttpContext.Response.Headers.TryAdd(header.Key, new StringValues(header.Value.ToArray()));
                }

                await stream.CopyToAsync(context.HttpContext.Response.Body);
                await context.HttpContext.Response.Body.FlushAsync();
            }
        }

        public static implicit operator HttpResponseMessage(HttpResponseMessageResult result)
        {
            return result.ResponseMessage;
        }

        public static implicit operator HttpResponseMessageResult(HttpResponseMessage message)
        {
            return new HttpResponseMessageResult(message);
        }
    }
}
