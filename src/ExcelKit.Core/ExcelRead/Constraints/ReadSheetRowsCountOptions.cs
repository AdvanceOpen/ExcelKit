using System;
using System.Text;
using System.Collections.Generic;
using ExcelKit.Core.Constraint.Enums;

namespace ExcelKit.Core.ExcelRead
{
	/// <summary>
	/// 读取Sheet行数量
	/// </summary>
	public class ReadSheetRowsCountOptions
	{
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
		/// 是否将空行算入总行数
		/// </summary>
		public bool ContainsEmptyRow { get; set; }

		/// <summary>
		/// 当直接传入Stream时，是否释放流(如果是指定的FilePath，则此选项不生效)
		/// 比如读取表头后，再根据表头读取具体内容，此时可指定为false不释放
		/// </summary>
		public bool IsDisposeStream { get; set; } = true;
	}
}
