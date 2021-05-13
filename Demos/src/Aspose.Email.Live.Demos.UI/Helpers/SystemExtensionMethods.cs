using Aspose.Email.Live.Demos.UI.Helpers;
using Aspose.Email.Live.Demos.UI.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System.Collections;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace System
{
    public static class SystemExtensionMethods
    {
		/// <summary>
		/// Return index of element
		/// </summary>
		/// <typeparam name="T">Array type</typeparam>
		/// <param name="self">Array reference</param>
		/// <param name="element">Element</param>
		/// <returns>Index or -1</returns>
		public static int IndexOf<T>(this T[] self, T element)
		{
			return Array.IndexOf<T>(self, element);
		}

		public static Response WithTrace(this Response self, IConfiguration config, Exception ex)
		{
			if (config.GetValue<bool?>("TraceEnabled") ?? false)
				self.Trace = ex.StackTrace;
			return self;
		}
		public static bool HasValue(this string value)
		{
			return !string.IsNullOrEmpty(value);
		}

		public static Stream AsStream(this string self)
		{
			return new MemoryStream(Encoding.UTF8.GetBytes(self));
		}

		/// <summary>
		/// Return true, if enumerable is null or empty
		/// </summary>
		/// <param name="self">Self reference</param>
		/// <returns>Bool</returns>
		public static bool IsNullOrEmpty(this IEnumerable self)
		{
			return self == null || !self.Cast<object>().Any();
		}

		/// <summary>
		/// Return true, if collection is null or empty
		/// </summary>
		/// <param name="self">Self reference</param>
		/// <returns>Bool</returns>
		public static bool IsNullOrEmpty(this ICollection self)
		{
			return self == null || self.Count == 0;
		}

		/// <summary>
		/// Return true, if string is null or empty
		/// </summary>
		/// <param name="self">Self reference</param>
		/// <returns>Bool</returns>
		public static bool IsNullOrEmpty(this string self)
		{
			return string.IsNullOrEmpty(self);
		}

		/// <summary>
		/// Return true, if string is null or contains only whitespaces
		/// </summary>
		/// <param name="self">Self reference</param>
		/// <returns>Bool</returns>
		public static bool IsNullOrWhiteSpace(this string self)
		{
			return string.IsNullOrWhiteSpace(self);
		}

		public static bool ParseToBoolOrDefault(this string self, bool defaultValue = false)
		{
			if (self.IsNullOrWhiteSpace())
				return defaultValue;

			bool.TryParse(self, out defaultValue);
			return defaultValue;
		}

		public static Task<T> ToTaskResult<T>(this T self)
		{
			return Task.FromResult(self);
		}

		public static byte[] ToArray(this Stream self)
		{
			if (self is MemoryStream stream)
				return stream.ToArray();
			return self.CopyToMemoryStream().ToArray();
		}

		public static MemoryStream CopyToMemoryStream(this Stream self)
		{
			var ms = new MemoryStream();
			self.CopyTo(ms);
			ms.Position = 0;
			return ms;
		}

		public static T ReadAs<T>(this Stream self)
		{
			var serializer = new JsonSerializer();

			using (var sr = new StreamReader(self))
			using (var jsonTextReader = new JsonTextReader(sr))
				return serializer.Deserialize<T>(jsonTextReader);
		}

		/// <summary>
		/// Reads HttpContent as json object with T type
		/// </summary>
		/// <typeparam name="T">Type for deserialization</typeparam>
		/// <param name="content">HttpContent to read</param>
		/// <returns>Async reading task</returns>
		public static async Task<T> ReadAsJsonAsync<T>(this HttpContent content)
		{
			using (var stream = await content.ReadAsStreamAsync())
			{
				using (var streamReader = new StreamReader(stream))
				using (var jsonReader = new JsonTextReader(streamReader))
					return new JsonSerializer().Deserialize<T>(jsonReader);
			}
		}

		public static IActionResult ToActionResult(this HttpResponseMessage message)
		{
			return new HttpResponseMessageResult(message);
		}

		/// <summary>
		/// Validate string of Regex
		/// </summary>
		public static bool IsValidRegex(this string pattern)
		{
			try
			{
				Regex.Match("", pattern);
			}
			catch (ArgumentException)
			{
				return false;
			}

			return true;
		}
	}
}
