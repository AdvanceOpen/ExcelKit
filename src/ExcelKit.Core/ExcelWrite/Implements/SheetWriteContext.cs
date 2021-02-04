using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using ExcelKit.Core.Helpers;
using ExcelKit.Core.Infrastructure;
using NPOI.SS.UserModel;

namespace ExcelKit.Core.ExcelWrite
{
	internal class SheetWriteContext : ISheetWriteContext
	{
		private ISheet _sheet;
		private string _fileId;
		private readonly uint _autoSplit;

		public SheetWriteContext(string fileId, uint autoSplit, ISheet sheet)
		{
			Inspector.NotNull(sheet, "sheet对象不能为空");
			Inspector.NotNullOrWhiteSpace(fileId, "文件标识不能为空");
			_sheet = sheet;
			_fileId = fileId;
			_autoSplit = autoSplit;
		}

		public uint AutoSplit => _autoSplit;
		public string SheetName => _sheet?.SheetName;

		public void AppendData<T>(string sheetName, T rowData) where T : class, new()
		{
			if (rowData == null)
				return;

			MultiStageExporter.AppendData(_fileId, sheetName, rowData, _autoSplit);
		}

		public void AppendData(string sheetName, Dictionary<string, object> rowData)
		{
			if (rowData == null || rowData.Count == 0)
				return;

			MultiStageExporter.AppendData(_fileId, sheetName, rowData, _autoSplit);
		}
	}
}
