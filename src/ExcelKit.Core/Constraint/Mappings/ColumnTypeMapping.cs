using System;
using System.Text;
using System.Collections.Generic;
using ExcelKit.Core.Constraint.Enums;

namespace ExcelKit.Core.Constraint.Mappings
{
	/// <summary>
	/// Excel中列的类型转换
	/// </summary>
	internal class ColumnTypeMapping
	{
		public static object Convert(string convertValue, ColumnType columnType, bool allowNull)
		{
			object result = null;
			switch (columnType)
			{
				case ColumnType.Int:
					result = allowNull && string.IsNullOrWhiteSpace(convertValue) ? 0 : int.Parse(convertValue);
					break;
				case ColumnType.NullInt:
					break;
				case ColumnType.Long:
					result = allowNull && string.IsNullOrWhiteSpace(convertValue) ? 0 : long.Parse(convertValue);
					break;
				case ColumnType.NullLong:
					break;
				case ColumnType.Decimal:
					//这样写主要是为了解决读取出1.0133E-2这种数据，这样才能转换
					result = allowNull && string.IsNullOrWhiteSpace(convertValue) ? 0 : System.Convert.ToDecimal(System.Convert.ToDouble(convertValue));
					break;
				case ColumnType.NullDecimal:
					break;
				case ColumnType.Time:
					var status = DateTime.TryParse(convertValue, out DateTime dateTime);
					if (allowNull && string.IsNullOrWhiteSpace(convertValue))
						result = DateTime.MinValue;
					else
						result = status ? dateTime : DateTime.FromOADate(System.Convert.ToDouble(convertValue));
					break;
				case ColumnType.NullTime:
					break;
				default:
					result = convertValue;
					break;
			}
			return result;
		}
	}
}
