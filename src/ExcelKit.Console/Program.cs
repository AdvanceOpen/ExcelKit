using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelKit.Consoles.Methods;

namespace ExcelKit.Consoles
{
	class Program
	{
		static async Task Main(string[] args)
		{
			//泛型导出
			ExcelWriteTest.GenericWrite();

			//动态导出
			//ExcelWriteTest.DynamicWrite();

			//泛型读取
			//ExcelReadTest.ReadSheetGeneric();
			Console.Read();
		}
	}
}
