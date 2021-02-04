using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelKit.Core.ExcelWrite
{
	/// <summary>
	/// 输出的Excel文件信息
	/// </summary>
	public class OutExcelInfo
	{
		/// <summary>
		/// 文件名
		/// </summary>
		public string FileName { get; set; }

		/// <summary>
		/// web下载时的ContentType
		/// </summary>
		public string WebContentType => "application/ms-excel";

		/// <summary>
		/// 文件大小，KB
		/// </summary>
		public long FileSizeKB => Stream == null ? 0 : Stream.Length / 1024L;

		/// <summary>
		/// 生成的Excel文件流
		/// </summary>
		public Stream Stream { get; set; }
	}
}
