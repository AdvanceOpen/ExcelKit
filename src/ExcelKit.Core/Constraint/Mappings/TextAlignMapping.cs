using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;
using ExcelKit.Core.Constraint.Enums;

namespace ExcelKit.Core.Constraint.Mappings
{
	/// <summary>
	/// 文本对齐方式映射(自定义->NPOI)
	/// </summary>
	internal class TextAlignMapping
	{
		public static HorizontalAlignment MapAlign(TextAlign textAlign)
		{
			HorizontalAlignment horizontalAlignment = HorizontalAlignment.Left;

			switch (textAlign)
			{
				case TextAlign.Left:
					horizontalAlignment = HorizontalAlignment.Left;
					break;
				case TextAlign.Center:
					horizontalAlignment = HorizontalAlignment.Center;
					break;
				case TextAlign.Right:
					horizontalAlignment = HorizontalAlignment.Right;
					break;
			}
			return horizontalAlignment;
		}
	}
}
