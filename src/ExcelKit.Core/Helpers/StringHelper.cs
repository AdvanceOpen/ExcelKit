using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelKit.Core.Helpers
{
	/// <summary>
	/// 字符串工具类
	/// </summary>
	internal class StringHelper
	{
		/// <summary>
		/// 是否是数字
		/// </summary>
		/// <param name="str">待检测的字符串</param>
		/// <returns></returns>
		public static bool IsNumber(string str)
		{
			return System.Text.RegularExpressions.Regex.IsMatch(str, @"^(\d+)$");
		}

		/// <summary>
		/// 去掉字符串中的数字
		/// </summary>
		/// <param name="key"></param>
		/// <returns></returns>
		public static string RemoveNumber(string key)
		{
			return System.Text.RegularExpressions.Regex.Replace(key, @"\d", "");
		}

		/// <summary>
		/// 去掉字符串中的非数字
		/// </summary>
		/// <param name="key"></param>
		/// <returns></returns>
		public static string RemoveNotNumber(string key)
		{
			return System.Text.RegularExpressions.Regex.Replace(key, @"[^\d]*", "");
		}
	}
}