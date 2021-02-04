using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelKit.Core.Infrastructure.Converter
{
	/// <summary>
	/// 小数点保留位数转换器
	/// </summary>
	/// <typeparam name="T">数据类型类型</typeparam>
	public class DecimalPointDigitConverter : IExportConverter<decimal, int>
	{
		/// <summary>
		/// 转换
		/// </summary>
		/// <param name="value">数值</param>
		/// <param name="digit">位数</param>
		/// <returns></returns>
		public object Convert(decimal value, int digit)
		{
			return System.Convert.ToDecimal(System.Convert.ToDouble(value).ToString($"f{digit}"));
		}
	}
}