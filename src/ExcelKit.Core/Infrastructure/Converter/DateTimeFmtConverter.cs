using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelKit.Core.Infrastructure.Converter
{
	/// <summary>
	/// 日期格式化类型转换器
	/// </summary>
	/// <typeparam name="T">数据类型类型</typeparam>
	public class DateTimeFmtConverter : IExportConverter<DateTime?, string>
	{
		public object Convert(DateTime? datetime, string format)
		{
			if (datetime == null)
				return "";
			if (string.IsNullOrWhiteSpace(format))
				return datetime.Value.ToString("yyyy-MM-dd HH:mm:ss");

			return datetime.Value.ToString(format);
		}
	}
}