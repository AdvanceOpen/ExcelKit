using System;
using System.Text;
using System.Collections.Generic;
using ExcelKit.Core.Constraint.Enums;

namespace ExcelKit.Core.ExcelRead
{
	/// <summary>
	/// 读取行
	/// </summary>
	public class ReadRowsOptions
	{
		/// <summary>
		/// 数据起始行（从1开始计算）
		/// </summary>
		public uint DataStartRow { get; set; } = 1;

		/// <summary>
		/// 数据结束行（从1开始计算）
		/// </summary>
		public uint? DataEndRow { get; set; }

		/// <summary>
		/// 读取方式(默认按照Sheet索引读)
		/// </summary>
		public ReadWay ReadWay { get; set; } = ReadWay.SheetIndex;

		/// <summary>
		/// Sheet索引
		/// </summary>
		public ushort SheetIndex { get; set; } = 1;

		/// <summary>
		/// Sheet名称
		/// </summary>
		public string SheetName { get; set; } = "Sheet1";

		/// <summary>
		/// 是否读取空单元格数据（要读取则需指定读取哪些列）
		/// </summary>
		public bool ReadEmptyCell { get; set; }

		/// <summary>
		/// ReadEmptyCell为true时，需指定此列头(A、B、C等)
		/// </summary>
		public string[] ColumnHeaders { get; set; }

		/// <summary>
		/// 行数据处理函数
		/// </summary>
		public Action<IList<string>> RowData { get; set; }

		/// <summary>
		/// 当直接传入Stream时，是否释放流(如果是指定的FilePath，则此选项不生效)
		/// 比如读取表头后，再根据表头读取具体内容，此时可指定为false不释放
		/// </summary>
		public bool IsDisposeStream { get; set; } = true;
	}
}
