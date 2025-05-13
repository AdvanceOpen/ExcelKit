using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelKit.Core.Attributes;
using ExcelKit.Core.Helpers;
using ExcelKit.Core.Infrastructure;
using ExcelKit.Core.Infrastructure.Exceptions;
using ExcelKit.Core.Infrastructure.Reflections;
using NPOI.SS.Formula.Functions;

namespace ExcelKit.Core.ExcelWrite
{
	/// <summary>
	/// ExcelContext
	/// </summary>
	internal class ExcelWriteContext : IExcelWriteContext
	{
		string _fileId;

		public ExcelWriteContext(string fileName)
		{
			_fileId = MultiStageExporter.CreateExcel(fileName);
		}

		public string GetSafeSheetName(string sheetName)
		{
			return MultiStageExporter.GetSafeSheetName(sheetName);
		}

		public ISheetWriteContext CrateSheet<T>(string sheetName, uint autoSplit = 1048200) where T : class, new()
		{
			var sortedProps = ReflectionHelper.NewInstance.GetSortedExportProps<T>();
			var sheet = MultiStageExporter.CreteSheet(_fileId, sheetName, sortedProps.Select(t => t.attr).ToList());
			return new SheetWriteContext(_fileId, autoSplit, sheet);
		}

		public ISheetWriteContext CrateSheet(string sheetName, List<ExcelKitAttribute> headers, uint autoSplit = 1048200)
		{
			Inspector.NotNullAndHasElement(headers, "动态表头信息不能为空");
			Inspector.Validation(headers.Count(t => string.IsNullOrWhiteSpace(t.Code)) > 0, "表头信息中存在Code为空的列字段");
			Inspector.Validation(headers.Count(t => string.IsNullOrWhiteSpace(t.Desc)) > 0, "表头信息中存在Desc为空的列字段");
			headers = headers.Where(t => t.IsIgnore == false).ToList();
			headers.RemoveAll(t => t.IsOnlyIgnoreWrite);

			var sheet = MultiStageExporter.CreteSheet(_fileId, sheetName, headers);
			return new SheetWriteContext(_fileId, autoSplit, sheet);
		}

		public OutExcelInfo Generate()
		{
			return MultiStageExporter.Generate(_fileId);
		}

		public string Save(string saveForder = null)
		{
			return MultiStageExporter.Save(_fileId, saveForder);
		}

		public void Dispose()
		{
			Dispose(true);
		}

		~ExcelWriteContext()
		{
			Dispose(false);
		}

		protected virtual void Dispose(bool disposing)
		{
			if (!disposing)
			{
				return;
			}

			MultiStageExporter.Dispose(_fileId);
			GC.SuppressFinalize(this);
		}
	}
}
