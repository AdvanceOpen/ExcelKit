using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelKit.Core.ExcelWrite
{
	internal class ExcelContextInfo
	{
		/// <summary>
		/// 文件标识
		/// </summary>
		public string FileId { get; set; }

		/// <summary>
		/// 工作簿
		/// </summary>
		public IWorkbook Workbook { get; set; }

		/// <summary>
		/// Sheet数量
		/// </summary>
		public int SheetCount { get; set; }

		/// <summary>
		/// Sheet信息(更多情况是用于内部自动拆分Sheet记录Sheet信息)
		/// </summary>
		public List<InnerSheetInfo> SheetInfo { get; set; } = new List<InnerSheetInfo>();

		/// <summary>
		/// 构建时间
		/// </summary>
		public DateTime BuildTime { get; set; }

		/// <summary>
		/// 导出文件名称
		/// </summary>
		public string ExportFileName { get; set; }
	}
}
