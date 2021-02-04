using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelKit.Core.Infrastructure.Converter
{
	/// <summary>
	/// 导出转换接口
	/// </summary>
	public interface IExportConverter<T1>// where T : class, new()
	{
		/// <summary>
		/// 转换
		/// </summary>
		/// <param name="obj">数据本身</param>
		/// <returns></returns>
		string Convert(T1 obj);
	}

	/// <summary>
	/// 导出转换接口
	/// </summary>
	public interface IExportConverter<T1, T2>
	{
		/// <summary>
		/// 转换
		/// </summary>
		/// <param name="obj1">默认为数据本身</param>
		/// <param name="obj2">默认为转换参数</param>
		/// <returns></returns>
		object Convert(T1 obj1, T2 obj2);
	}
}
