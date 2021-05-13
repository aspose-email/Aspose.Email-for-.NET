using Microsoft.Extensions.Logging;
using System;
using System.Runtime.CompilerServices;

namespace Aspose.Email.Live.Demos.UI
{
    public static class LoggerExtensionMethods
    {
		public static void LogEmailError(this ILogger logger, Exception ex, string appOrServiceName,
			[CallerFilePath] string callerFilePath = "",
			[CallerMemberName] string callerMemberName = "",
			[CallerLineNumber] int sourceLineNumber = 0)
		{
			var logMsg = $@"CallerName = {callerFilePath}, MethodName = {callerMemberName}, Line = {sourceLineNumber}";
			logger.LogError(ex, logMsg, appOrServiceName);
		}
	}
}
