using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Net.Http;

namespace Aspose.Email.Live.Demos.UI.Helpers
{
    /// <summary>
    /// Default Exception Filter Attribute
    /// </summary>
    public class DefaultExceptionFilterAttribute : ExceptionFilterAttribute
	{
        public override void OnException(ExceptionContext context)
        {
			var actionName = context.ActionDescriptor.DisplayName;

            Startup.ExceptionFilterLogger.LogError(context.Exception, $"{actionName} error", actionName, "");

            var resp = new HttpResponseMessageResult(new HttpResponseMessage(System.Net.HttpStatusCode.InternalServerError)
            {
                ReasonPhrase = actionName,
            });

            if (Startup.Configuration.GetValue<bool?>("TraceEnabled") ?? false)
                context.HttpContext.Response.Body = context.Exception.StackTrace.AsStream();

            context.Result = resp;
        }
    }
}
