using System;
using System.Text;
using System.Collections.Generic;
using ExcelKit.Core.Constraint.Enums;

namespace ExcelKit.Core.ExcelRead
{
	/// <summary>
	/// 读取Sheet
	/// </summary>
	public class ReadSheetOptions<T> where T : class, new()
	{
		/// <summary>
		/// 表头数据所在行
		/// </summary>
		public uint HeadRow { get; set; } = 1;

		/// <summary>
		/// 数据起始行（从1开始计算）
		/// </summary>
		public uint DataStartRow { get; set; } = 2;

		/// <summary>
		/// 数据结束行（从1开始计算）
		/// </summary>
		public uint? DataEndRow { get; set; }

		/// 读取方式(默认按照Sheet索引读)
		/// </summary>
		public ReadWay ReadWay { get; set; } = ReadWay.SheetIndex;

		/// <summary>
		/// Sheet名称，默认Sheet1
		/// </summary>
		public string SheetName { get; set; } = "Sheet1";

		/// <summary>
		/// Sheet索引
		/// </summary>
		public ushort SheetIndex { get; set; } = 1;

		/// <summary>
		/// 读取成功的行数据处理函数(Param1:读取出的数据  Param2：行索引)
		/// </summary>
		public Action<T, uint> SucData { get; set; }

		/// <summary>
		/// 读取失败的行数据数据处理函数【失败的原始数据，失败数据信息汇总】
		/// </summary>
		public Action<Dictionary<string, object>, List<(string rowIndex, string columnName, string cellValue, string errorMsg)>> FailData { get; set; }
	}
}
