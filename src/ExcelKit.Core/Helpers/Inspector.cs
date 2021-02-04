using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelKit.Core.Infrastructure.Exceptions;

namespace ExcelKit.Core.Helpers
{
	/// <summary>
	/// 校验辅助类
	/// </summary>
	public class Inspector
	{
		/// <summary>
		/// 值不为NULL
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="argument">要判断的值</param>
		/// <param name="tipMsg">提示信息</param>
		public static void NotNull<T>(T argument, string tipMsg) where T : class
		{
			argument = argument ?? throw new ExcelKitException(tipMsg);
		}

		/// <summary>
		/// 如果不为空
		/// </summary>
		/// <param name="param"></param>
		/// <param name="tipmsg"></param>
		public static void IfNotNull(object param, string tipmsg)
		{
			if (param != null)
			{
				throw new ExcelKitException(tipmsg);
			}
		}

		/// <summary>
		/// 校验器
		/// </summary>
		/// <param name="condition"></param>
		/// <param name="tipmsg"></param>
		public static void Validation(bool condition, string tipmsg)
		{
			if (condition)
			{
				throw new ExcelKitException(tipmsg);
			}
		}

		/// <summary>
		/// 值不为NULL且必须包含元素
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="arguments">要判断的值</param>
		/// <param name="tipMsg">提示信息</param>
		public static void NotNullAndHasElement<T>(IEnumerable<T> arguments, string tipMsg) where T : class
		{
			if (arguments == null || arguments.Count() == 0)
				throw new ExcelKitException(tipMsg);
		}

		/// <summary>
		/// 值不为NULL且不为空
		/// </summary>
		/// <param name="argument">要判断的值</param>
		/// <param name="tipMsg">提示信息</param>
		public static void NotNullOrWhiteSpace(string argument, string tipMsg)
		{
			if (string.IsNullOrWhiteSpace(argument))
				throw new ExcelKitException(tipMsg);
		}

		/// <summary>
		/// 值必须在两者之间
		/// </summary>
		/// <param name="value">要判断的值</param>
		/// <param name="min">最小值</param>
		/// <param name="max">最大值</param>
		/// <param name="tipMsg">提示信息</param>
		public static void Between<T>(T value, T min, T max, string tipMsg) where T : IComparable<T>
		{
			if (value.CompareTo(min) < 0 || value.CompareTo(max) > 0)
			{
				throw new ExcelKitException(tipMsg);
			}
		}

		/// <summary>
		/// 值必须大于
		/// </summary>
		/// <param name="value">要判断的值</param>
		/// <param name="compareValue">比较的值</param>
		/// <param name="tipMsg">提示信息</param>
		public static void MoreThan<T>(T value, T compareValue, string tipMsg) where T : IComparable<T>
		{
			if (value.CompareTo(compareValue) <= 0)
			{
				throw new ExcelKitException(tipMsg);
			}
		}

		/// <summary>
		/// 值必须大于等于
		/// </summary>
		/// <param name="value">要判断的值</param>
		/// <param name="compareValue">比较的值</param>
		/// <param name="tipMsg">提示信息</param>
		public static void MoreThanOrEqual<T>(T value, T compareValue, string tipMsg) where T : IComparable<T>
		{
			if (value.CompareTo(compareValue) < 0)
			{
				throw new ExcelKitException(tipMsg);
			}
		}

		/// <summary>
		/// 值必须小于
		/// </summary>
		/// <param name="value">要判断的值</param>
		/// <param name="compareValue">比较的值</param>
		/// <param name="tipMsg">提示信息</param>
		public static void LessThan<T>(T value, T compareValue, string tipMsg) where T : IComparable<T>
		{
			if (value.CompareTo(compareValue) >= 0)
			{
				throw new ExcelKitException(tipMsg);
			}
		}

		/// <summary>
		/// 值必须小于等于
		/// </summary>
		/// <param name="value">要判断的值</param>
		/// <param name="compareValue">比较的值</param>
		/// <param name="tipMsg">提示信息</param>
		public static void LessThanOrEqual<T>(T value, T compareValue, string tipMsg) where T : IComparable<T>
		{
			if (value.CompareTo(compareValue) > 0)
			{
				throw new ExcelKitException(tipMsg);
			}
		}
	}
}
