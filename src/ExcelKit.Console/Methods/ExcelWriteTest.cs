using System;
using System.Text;
using System.Collections.Generic;
using ExcelKit.Core.Attributes;
using ExcelKit.Core.ExcelWrite;
using ExcelKit.Core.Infrastructure.Factorys;
using System.Threading.Tasks;
using System.Diagnostics;

namespace ExcelKit.Consoles.Methods
{
	public class ExcelWriteTest
	{
		public static string GenericWrite(bool isOpen = true)
		{
			string filePath;
			using (var context = ContextFactory.GetWriteContext("测试导出文件"))
			{
				var sheet = context.CrateSheet<UserDto>($"Sheet1");
				for (int index = 0; index < 100; index++)
				{
					bool? sex = null;
					if (index != 0)
						sex = index % 2 == 0;

					sheet.AppendData<UserDto>($"Sheet1", new UserDto { Account = $"{index}-2010211", Name = $"{index}-用户用户", CreateDate = DateTime.Now, Sex = sex });
				}
				//var sw = new Stopwatch();
				//sw.Start();

				//Parallel.For(1, 4, index =>
				//{
				//	var sheet = context.CrateSheet<UserDto>($"Sheet{index}");

				//	for (int i = 0; i < 10000; i++)
				//	{
				//		sheet.AppendData<UserDto>($"Sheet{index}", new UserDto { Account = $"{index}-{i}-2010211", Name = $"{index}-{i}-用户用户" });
				//	}
				//});
				//var millsec = sw.ElapsedMilliseconds;
				//Console.WriteLine("共计耗时：" + millsec);

				filePath = context.Save();
				Console.WriteLine($"文件路径：{filePath}");
			}

			//浏览文件
			if (isOpen)
			{
				System.Diagnostics.Process.Start("explorer.exe", filePath);
			}

			return filePath;
		}

		public static string DynamicWrite()
		{
			string filePath;
			using (var context = ContextFactory.GetWriteContext("测试导出文件"))
			{
				//注意CreateSheet方法最后一个字段，指定多少条数据自动拆分一个新Sheet，不指定默认为单Sheet最大数据量1048200
				var sheet = context.CrateSheet("Sheet1", new List<ExcelKitAttribute>()
				{
					new ExcelKitAttribute(){ Code = "Account", Desc = "账号",Width = 60 },
					new ExcelKitAttribute(){ Code = "Name", Desc = "昵称" }
				}, 1048200);

				for (int i = 0; i < 104; i++)
				{
					sheet.AppendData("Sheet1", new Dictionary<string, object>()
					{
						{"Account", $"{i}-2010211" }, {"Name", $"{i}-用户用户" }
					});
				}

				filePath = context.Save();
				Console.WriteLine($"文件路径：{filePath}");
			}

			//浏览文件
			System.Diagnostics.Process.Start("explorer.exe", filePath);

			return filePath;
		}
	}
}
