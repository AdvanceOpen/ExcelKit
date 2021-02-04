using System;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using ExcelKit.Core.ExcelWrite;
using ExcelKit.Core.ExcelRead;
using Newtonsoft.Json;
using ExcelKit.Core.Constraint.Enums;
using ExcelKit.Core.Infrastructure.Factorys;

namespace ExcelKit.Consoles.Methods
{
	public class ExcelReadTest
	{
		public static void SheetIndexReadRows()
		{
			var context = ContextFactory.GetReadContext();
			context.ReadRows("测试导出文件.xlsx", new ReadRowsOptions()
			{
				RowData = rowdata =>
				{
					Console.WriteLine(JsonConvert.SerializeObject(rowdata));
				}
			});
		}

		public static void SheetNameReadRows()
		{
			var context = ContextFactory.GetReadContext();
			context.ReadRows("测试导出文件.xlsx", new ReadRowsOptions()
			{
				ReadWay = ReadWay.SheetName,
				RowData = rowdata =>
				{
					Console.WriteLine(JsonConvert.SerializeObject(rowdata));
				}
			});
		}

		public static void ReadSheetGeneric()
		{
			var context = ContextFactory.GetReadContext();
			context.ReadSheet("测试导出文件.xlsx", new ReadSheetOptions<UserDto>()
			{
				SucData = (rowdata, rowindex) =>
				{
					Console.WriteLine(JsonConvert.SerializeObject(rowdata));
				},
				FailData = (odata, failinfo) =>
				{
					Console.WriteLine($"读取失败，{failinfo.FirstOrDefault().errorMsg}");
				}
			});
		}

		public static void ReadSheetDic()
		{
			var context = ContextFactory.GetReadContext();
			context.ReadSheet("测试导出文件.xlsx", new ReadSheetDicOptions()
			{
				DataEndRow = 10,
				ExcelFields = new (string field, ColumnType type, bool allowNull)[]
				{
					("账号",ColumnType.String,false),("昵称",ColumnType.String,false)
				},
				SucData = (rowdata, rowindex) =>
				{
					Console.WriteLine(JsonConvert.SerializeObject(rowdata));
				},
				FailData = (odata, failinfo) =>
				{
					Console.WriteLine($"读取失败，{failinfo.FirstOrDefault().errorMsg}");
				}
			});
		}
	}
}
