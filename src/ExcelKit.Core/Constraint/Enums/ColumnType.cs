using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace ExcelKit.Core.Constraint.Enums
{
	/// <summary>
	/// Excel中列的类型
	/// </summary>
	public enum ColumnType
	{
		String,
		Int,
		NullInt,
		Long,
		NullLong,
		Decimal,
		NullDecimal,
		Time,
		NullTime
	}
}
