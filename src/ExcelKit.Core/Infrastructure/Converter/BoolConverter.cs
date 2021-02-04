using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelKit.Core.Infrastructure.Converter
{
	/// <summary>
	/// bool类型转换器
	/// </summary>
	/// <remarks>
	/// param1：数据本身   param2：自定义显示的文字
	/// true：默认为"是"     false：默认为"否"
	/// </remarks>
	public class BoolConverter : IExportConverter<bool, string>, IExportConverter<bool?, string>
	{
		public object Convert(bool obj1, string obj2)
		{
			string[] arrs = this.Split(obj2);

			return obj1 switch
			{
				true => arrs[0],
				false => arrs[1],
			};
		}

		public object Convert(bool? obj1, string obj2)
		{
			string[] arrs = this.Split(obj2);

			return obj1 switch
			{
				true => arrs[0],
				false => arrs[1],
				_ => string.Empty
			};
		}

		private string[] Split(string obj2)
		{
			if (string.IsNullOrEmpty(obj2))
				return new string[] { "是", "否" };

			return obj2.Split('|').Where(t => !string.IsNullOrWhiteSpace(t)).ToArray();
		}
	}
}
