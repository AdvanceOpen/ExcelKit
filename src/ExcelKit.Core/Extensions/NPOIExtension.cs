using System;
using System.Collections.Generic;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelKit.Core.Extensions
{
	public class NPOIExtension
	{
		/// <summary>
		/// 复制行
		/// </summary>
		/// <param name="sheet">所在Sheet</param>
		/// <param name="sourceRow">被复制的行</param>
		/// <param name="insertRow">插入行索引</param>
		/// <param name="insertRowCount">插入行数量</param>
		public static void CopyRows(ISheet sheet, IRow sourceRow, int insertRow, int insertRowCount)
		{
			//批量移动行（--开始行   --结束行   移动大小(行数)--往下移动   是否复制行高   是否重置行高）
			sheet.ShiftRows(insertRow, sheet.LastRowNum, insertRowCount, true, false);

			//对批量移动后空出的空行插，创建相应的行
			for (int i = insertRow; i < insertRow + insertRowCount; i++)
			{
				IRow targetRow = sheet.CreateRow(i);
				ICell sourceCell = null;
				ICell targetCell = null;

				for (int m = sourceRow.FirstCellNum; m < sourceRow.LastCellNum; m++)
				{
					sourceCell = sourceRow.GetCell(m);
					if (sourceCell == null)
						continue;

					targetCell = targetRow.CreateCell(m);
					//targetCell.Encoding = sourceCell.Encoding;

					targetCell.CellStyle = sourceCell.CellStyle;
					targetCell.SetCellType(sourceCell.CellType);
				}
			}
		}

		/// <summary>
		/// 根据Excel列类型获取列的值(一般用于读取)
		/// </summary>
		/// <param name="cell">Excel列</param>
		/// <returns></returns>
		public static string GetCellValue(ICell cell)
		{
			if (cell == null)
				return string.Empty;
			if (cell.CellType.ToString() == "System.DBNull")
			{
				return string.Empty;
			}
			switch (cell.CellType)
			{
				case CellType.Blank:
					return string.Empty;
				case CellType.Boolean:
					return cell.BooleanCellValue.ToString();
				case CellType.Error:
					return cell.ErrorCellValue.ToString();
				case CellType.Numeric:
					//日期类型
					return HSSFDateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue.ToString() : cell.NumericCellValue.ToString();
				case CellType.String:
					return cell.StringCellValue;
				case CellType.Formula:
					try
					{
						HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
						e.EvaluateInCell(cell);
						return cell.ToString();
					}
					catch
					{
						return cell.NumericCellValue.ToString();
					}
				default:
					return cell.ToString();
			}
		}
	}
}
