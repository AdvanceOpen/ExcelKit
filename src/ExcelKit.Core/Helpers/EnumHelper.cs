using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace ExcelKit.Core.Helpers
{
	/// <summary>
	/// 枚举辅助类
	/// </summary>
	internal class EnumHelper
	{
		/// <summary>
		/// 枚举信息
		/// </summary>
		internal class EnumInfo
		{
			/// <summary>
			/// 枚举名称
			/// </summary>
			public string EnumName { get; set; }

			/// <summary>
			/// 枚举描述
			/// </summary>
			public string EnumDesc { get; set; }

			/// <summary>
			/// 枚举值
			/// </summary>
			public int EnumValue { get; set; }
		}

		/// <summary>
		/// 缓存的信息
		/// </summary>
		static ConcurrentDictionary<string, List<EnumInfo>> _cache = new ConcurrentDictionary<string, List<EnumInfo>>();

		/// <summary>
		/// 获取枚举信息
		/// </summary>
		/// <param name="enumType">枚举类型</param>
		/// <returns></returns>
		protected internal static List<EnumInfo> GetEnumInfo(Type enumType)
		{
			if (!enumType.IsEnum)
			{
				return null;
			}

			if (_cache.ContainsKey(enumType.AssemblyQualifiedName))
			{
				return _cache[enumType.AssemblyQualifiedName];
			}

			List<EnumInfo> enumInfos = new List<EnumInfo>();
			System.Reflection.FieldInfo[] fieldinfos = enumType.GetFields();

			foreach (System.Reflection.FieldInfo field in fieldinfos)
			{
				if (!field.FieldType.IsEnum) { continue; }
				var objs = field.GetCustomAttributes(typeof(DescriptionAttribute), false).Cast<DescriptionAttribute>();

				enumInfos.Add(new EnumInfo()
				{
					EnumName = field.Name,
					EnumDesc = objs == null || objs.Count() == 0 ? field.Name : objs.FirstOrDefault().Description?.Trim(),
					EnumValue = (int)field.GetValue(fieldinfos)
				});
			}
			_cache.TryAdd(enumType.AssemblyQualifiedName, enumInfos);

			return enumInfos;
		}
	}
}
