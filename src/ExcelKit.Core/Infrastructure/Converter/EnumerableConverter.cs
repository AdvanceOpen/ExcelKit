using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelKit.Core.Infrastructure.Converter
{
	/// <summary>
	/// 简单集合类型转换器
	/// </summary>
	/// <typeparam name="T">数据类型类型</typeparam>
	public class EnumerableConverter<T> : IExportConverter<IEnumerable<T>>
	{
		public string Convert(IEnumerable<T> obj)
		{
			return obj == null || obj.Count() == 0 ? string.Empty : string.Join("，", obj);
		}
	}
}
