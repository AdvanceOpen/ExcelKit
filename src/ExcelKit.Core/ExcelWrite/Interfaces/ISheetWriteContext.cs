using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelKit.Core.ExcelWrite
{
	/// <summary>
	/// Sheet上下文
	/// </summary>
	public interface ISheetWriteContext
	{
		/// <summary>
		/// 自动拆分Sheet的条数
		/// </summary>
		uint AutoSplit { get; }

		/// <summary>
		/// Sheet名称
		/// </summary>
		string SheetName { get; }

		/// <summary>
		/// 追加数据到Sheet中
		/// </summary>
		void AppendData<T>(string sheetName, T data) where T : class, new();

		/// <summary>
		/// 追加数据到Sheet中
		/// </summary>
		/// <param name="sheetName"></param>
		/// <param name="rowData">行数据(Key:字段名Code编码值  Value：要导出的数据)</param>
		void AppendData(string sheetName, Dictionary<string, object> rowData);
	}
}
