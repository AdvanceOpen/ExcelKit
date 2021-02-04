using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelKit.Core.Constraint.Mappings
{
	/// <summary>
	/// 数字类型转换为Decimal类型
	/// </summary>
	/// <remarks>为了兼容小数导出时保留的小数位</remarks>
	internal class NumToDecimalMapping
	{
		public static object IfNeedToDecimal(object value)
		{
			if (value is byte || value is int || value is long || value is short || value is float || value is double || value is decimal)
			{
				return Convert.ToDecimal(value);
			}
			return value;
		}
	}
}
