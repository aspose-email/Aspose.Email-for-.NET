using Newtonsoft.Json;
using System.Collections;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace System
{
	/// <summary>
	/// System namespace extension methods
	/// </summary>
	public static class SystemExtensionMethods
    {
		/// <summary>
		/// Reads HttpContent as json object with T type
		/// </summary>
		/// <typeparam name="T">Type for deserialization</typeparam>
		/// <param name="content">HttpContent to read</param>
		/// <returns>Async reading task</returns>
		public static Task<T> ReadAsJsonAsync<T>(this HttpContent content)
		{
			return content.ReadAsStreamAsync().ContinueWith(task =>
			{
				using (var streamReader = new StreamReader(task.Result))
					using (var jsonReader = new JsonTextReader(streamReader))
						return new JsonSerializer().Deserialize<T>(jsonReader);

			});
		}

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
	}
}
