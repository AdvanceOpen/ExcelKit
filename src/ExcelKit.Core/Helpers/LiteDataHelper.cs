using System;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using ExcelKit.Core.Attributes;
using ExcelKit.Core.ExcelWrite;
using ExcelKit.Core.ExcelRead;
using ExcelKit.Core.Infrastructure.Factorys;
using System.Text.RegularExpressions;

namespace ExcelKit.Core.Helpers
{
	/// <summary>
	/// 轻量级数据导出读取辅助类
	/// </summary>
	public class LiteDataHelper
	{
		#region 导出

		/// <summary>
		/// 设置导出的Excel名称
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="fileName"></param>
		/// <returns></returns>
		private static string SetFileName<T>(string fileName = null) where T : class, new()
		{
			if (string.IsNullOrWhiteSpace(fileName))
			{
				var classAttrs = typeof(T).GetCustomAttributes(false);
				if (classAttrs?.Any() == true)
				{
					var desc = ((ExcelKitAttribute)classAttrs.First()).Desc;
					fileName = string.IsNullOrWhiteSpace(desc) ? typeof(T).Name : desc;
				}
				fileName = string.IsNullOrWhiteSpace(fileName) ? typeof(T).Name : fileName;
			}
			return fileName;
		}

		/// <summary>
		/// 快速导出Excel（用于web下载，适用于轻量级数据）
		/// </summary>
		/// <typeparam name="T">泛型类</typeparam>
		/// <param name="data">数据集</param>
		/// <param name="sheetName">Sheet名称，默认Sheet1</param>
		/// <param name="fileName">文件名</param>
		/// <returns></returns>
		public static OutExcelInfo ExportToWebDown<T>(IList<T> data, string sheetName = "Sheet1", string fileName = null) where T : class, new()
		{
			fileName = SetFileName<T>(fileName);
			using (var context = ContextFactory.GetWriteContext(fileName))
			{
				var sheet = context.CrateSheet<T>(sheetName);
				foreach (var item in data)
				{
					sheet.AppendData<T>(sheetName, item);
				}

				return context.Generate();
			}
		}

		/// <summary>
		/// 快速导出Excel（直接保存到磁盘，适用于轻量级数据）
		/// </summary>
		/// <typeparam name="T">泛型类</typeparam>
		/// <param name="data">数据集</param>
		/// <param name="sheetName">Sheet名称，默认Sheet1</param>
		/// <param name="fileName">文件名</param>
		/// <returns></returns>
		public static string ExportToDisk<T>(IList<T> data, string sheetName = "Sheet1", string filePath = null, string fileName = null) where T : class, new()
		{
			fileName = SetFileName<T>(fileName);
			using (var context = ContextFactory.GetWriteContext(fileName))
			{
				var sheet = context.CrateSheet<T>(sheetName);
				foreach (var item in data)
				{
					sheet.AppendData<T>(sheetName, item);
				}

				return context.Save(filePath);
			}
		}

		#endregion

		#region 读取

		/// <summary>
		/// 读取Excel(轻量级数据)
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="filePath"></param>
		/// <returns></returns>
		public static (List<T> sucData, List<(Dictionary<string, object> odata, List<(string rowIndex, string columnName, string cellValue, string errorMsg)> failInfo)> failData) Read<T>(string filePath) where T : class, new()
		{
			var sucDatas = new List<T>();
			var failDatas = new List<(Dictionary<string, object> odata, List<(string rowIndex, string columnName, string cellValue, string errorMsg)> failInfo)>();

			ContextFactory.GetReadContext().ReadSheet("测试导出文件.xlsx", new ReadSheetOptions<T>()
			{
				SucData = (rowdata, rowindex) =>
				{
					sucDatas.Add(rowdata);
				},
				FailData = (odata, failinfo) =>
				{
					failDatas.Add((odata, failinfo));
				}
			});

			return (sucDatas, failDatas);
		}

		#endregion

		#region SheetName

		/// <summary>
		/// 获取安全的Sheet名称
		/// </summary>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		public static string GetSafeSheetName(string sheetName)
		{
			Inspector.NotNullOrWhiteSpace(sheetName, "Sheet名称不能为空");
			return Regex.Replace(sheetName, "[\\[\\]\\^[\\]\\/\\-_*×――(^)|'$%~!@#$…&%￥—+=<>《》!！??？:：•`·、。，；,.;\"‘’“”-]", "_");
		}

		/// <summary>
		/// 是否安全的Sheet名称
		/// </summary>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		public static bool IsSafeSheetName(string sheetName)
		{
			Inspector.NotNullOrWhiteSpace(sheetName, "Sheet名称不能为空");
			return Regex.IsMatch(sheetName, "[\\[\\]\\^[\\]\\/\\-_*×――(^)|'$%~!@#$…&%￥—+=<>《》!！??？:：•`·、。，；,.;\"‘’“”-]");
		}

		#endregion
	}
}
