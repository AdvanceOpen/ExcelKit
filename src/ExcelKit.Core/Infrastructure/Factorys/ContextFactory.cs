using System;
using System.Collections.Generic;
using System.Text;
using ExcelKit.Core.ExcelRead;
using ExcelKit.Core.ExcelWrite;

namespace ExcelKit.Core.Infrastructure.Factorys
{
	public static class ContextFactory
	{
		public static IReadExcelContext GetReadContext()
		{
			return new ReadExcelContext();
		}

		public static IExcelWriteContext GetWriteContext(string fileName)
		{
			return new ExcelWriteContext(fileName);
		}
	}
}
