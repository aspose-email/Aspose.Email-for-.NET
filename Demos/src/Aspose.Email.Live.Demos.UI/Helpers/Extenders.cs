using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Web;

namespace Aspose.Email.Live.Demos.UI.Helpers
{
	/// <summary>
	/// Extenders
	/// </summary>
	internal static class Extenders
	{
		/// <summary>
		/// Parses string as enum of specified type.
		/// </summary>
		/// <typeparam name="T">Enum type.</typeparam>
		/// <param name="source">Source string.</param>
		/// <param name="ignoreCase">Is operation case insensitive.</param>
		/// <returns>Enum value.</returns>
		public static T ParseEnum<T>(this string source, bool ignoreCase = true) where T : struct => (T)Enum.Parse(typeof(T), source, ignoreCase);


		public static string Format(this string str, params Expression<Func<string, object>>[] args)
		{
			var parameters = args.ToDictionary(
				key => string.Format("{{{0}}}", key.Parameters[0].Name),
				value => value.Compile()(value.Parameters[0].Name)
			);

			var sb = new StringBuilder(str);
			foreach (var kv in parameters)
				sb.Replace(kv.Key, kv.Value != null ? kv.Value.ToString() : "");

			return sb.ToString();
		}

		public static string FirstAlpha(this string value, params Expression<Func<string, object>>[] args) =>
			string.IsNullOrEmpty(value)
				? null
				: value[0].ToString();

		public static IEnumerable<string> Sliding(this string source, int window) =>
			source.Length < window
				? Enumerable.Empty<string>()
				: from i in Enumerable.Range(0, source.Length - (window - 1))
				  select source[i..][..window];
	}
}
