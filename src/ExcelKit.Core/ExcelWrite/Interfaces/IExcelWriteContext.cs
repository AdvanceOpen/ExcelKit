using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelKit.Core.Attributes;

namespace ExcelKit.Core.ExcelWrite
{
	/// <summary>
	/// Excel上下文
	/// </summary>
	public interface IExcelWriteContext : IDisposable
	{
		/// <summary>
		/// 安全的Sheet名称
		/// </summary>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		string SafeSheetName(string sheetName);

		/// <summary>
		/// 创建Sheet
		/// </summary>
		/// <param name="sheetName">Sheet名称</param>
		/// <param name="autoSplit">多少条后自动拆分Sheet</param>
		/// <returns></returns>
		ISheetWriteContext CrateSheet<T>(string sheetName, uint autoSplit = 1048200) where T : class, new();

		/// <summary>
		/// 创建Sheet
		/// </summary>
		/// <param name="sheetName">Sheet名称</param>
		/// <param name="headers">表头行字段信息</param>
		/// <param name="autoSplit">多少条后自动拆分Sheet</param>
		/// <returns></returns>
		ISheetWriteContext CrateSheet(string sheetName, List<ExcelKitAttribute> headers, uint autoSplit = 1048200);

		/// <summary>
		/// 生成输出的Excel信息
		/// </summary>
		OutExcelInfo Generate();

		/// <summary>
		/// 生成并保存Excel
		/// </summary>
		/// <param name="savePath">保存路径</param>
		/// <returns>文件路径</returns>
		string Save(string savePath = null);
	}
}
