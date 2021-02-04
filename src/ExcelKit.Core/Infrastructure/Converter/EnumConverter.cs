using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Linq;
using ExcelKit.Core.Helpers;

namespace ExcelKit.Core.Infrastructure.Converter
{
	/// <summary>
	/// 枚举类型转换器
	/// </summary>
	/// <typeparam name="EnumT"></typeparam>
	public class EnumConverter<EnumT> : IExportConverter<EnumT>
	{
		public string Convert(EnumT obj)
		{
			if (obj == null || !obj.GetType().IsEnum)
			{
				return string.Empty;
			}
			Type type = obj.GetType();
			return EnumHelper.GetEnumInfo(type)?.FirstOrDefault(t => t.EnumName == obj.ToString())?.EnumDesc ?? "";

			//下述方式不再使用，采用上述缓存的方式获取
			//MemberInfo[] memInfo = type.GetMember(obj.ToString());
			//if (memInfo != null && memInfo.Length > 0)
			//{
			//	object[] attrs = memInfo[0].GetCustomAttributes(typeof(System.ComponentModel.DescriptionAttribute), false);
			//	if (attrs != null && attrs.Length > 0)
			//		return ((System.ComponentModel.DescriptionAttribute)attrs[0]).Description;
			//}
			//return obj.ToString();
		}
	}
}
