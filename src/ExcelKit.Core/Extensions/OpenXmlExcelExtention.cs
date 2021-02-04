using System;
using System.Linq;
using System.Xml;
using System.Text;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelKit.Core.Extensions
{
	internal static class OpenXmlExcelExtention
	{
		/// <summary>
		/// 获取所有Sheet名称
		/// </summary>
		/// <param name="workbookPart"></param>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		public static IEnumerable<string> GetSheetNames(this WorkbookPart workbookPart)
		{
			return workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Select(t => t.Name.ToString()).ToList();
		}

		/// <summary>
		/// 获取Sheet
		/// </summary>
		/// <param name="workbookPart"></param>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		public static Sheet GetSheet(this WorkbookPart workbookPart, string sheetName)
		{
			//return workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(t => t.Name == sheetName);
			return workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().FirstOrDefault(t => t.Name == sheetName);
		}

		/// <summary>
		/// 获取行
		/// </summary>
		/// <param name="sheetData"></param>
		/// <param name="rowIndex"></param>
		/// <returns></returns>
		public static Row GetRow(this SheetData sheetData, uint rowIndex)
		{
			return sheetData.Elements<Row>().FirstOrDefault(t => t.RowIndex == rowIndex);
		}

		/// <summary>
		/// 获取单元格
		/// </summary>
		/// <param name="sheetData">sheet信息</param>
		/// <param name="columnName">列名称</param>
		/// <param name="rowIndex">行索引下标</param>
		/// <returns></returns>
		public static Cell GetCell(this SheetData sheetData, string columnName, uint rowIndex)
		{
			Row row = GetRow(sheetData, rowIndex);

			if (row == null)
			{
				return null;
			}

			return row.Elements<Cell>().FirstOrDefault(t => string.Compare(t.CellReference.Value, columnName + rowIndex, true) == 0);
		}

		/// <summary>
		/// 获取或创建单元格
		/// </summary>
		/// <param name="sheetData">sheet信息</param>
		/// <param name="columnName">列名称</param>
		/// <param name="rowIndex">行索引下标</param>
		/// <returns></returns>
		public static Cell GetOrCreateCell(this SheetData sheetData, string columnName, uint rowIndex)
		{
			string cellReference = columnName + rowIndex;

			Row row;
			//如果工作簿中不存在指定行则插入一行
			if (sheetData.Elements<Row>().Count(t => t.RowIndex == rowIndex) != 0)
			{
				row = sheetData.Elements<Row>().FirstOrDefault(t => t.RowIndex == rowIndex);
			}
			else
			{
				row = new Row() { RowIndex = rowIndex };
				sheetData.Append(row);
			}

			return row.GetOrCreateCell(cellReference);
		}

		/// <summary>
		/// 获取或创建单元格(通过行列标识，如A1，B2)
		/// </summary>
		/// <param name="sheetData">sheet信息</param>
		/// <param name="columnName">列名称</param>
		/// <param name="rowIndex">行索引下标</param>
		/// <returns></returns>
		public static Cell GetOrCreateCell(this Row row, string cellReference)
		{
			// If there is not a cell with the specified column name, insert one.  
			if (row.Elements<Cell>().Count(c => c?.CellReference?.Value == cellReference) > 0)
			{
				return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
			}
			else
			{
				// Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
				Cell refCell = null;
				foreach (Cell cell in row.Elements<Cell>())
				{
					if (cell.CellReference.Value.Length == cellReference.Length)
					{
						if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
						{
							refCell = cell;
							break;
						}
					}
				}

				Cell newCell = new Cell() { CellReference = cellReference };
				row.InsertBefore(newCell, refCell);
				return newCell;
			}
		}

		/// <summary>
		/// 获取单元格的值
		/// </summary>
		/// <param name="cell">单元格</param>
		/// <param name="shareStringPart"></param>
		/// <returns></returns>
		public static string GetValue(this Cell cell, SharedStringTablePart shareStringPart)
		{
			if (cell == null)
				return null;

			string cellvalue = cell.InnerText;
			if (cell.DataType != null)
			{
				if (cell.DataType == CellValues.SharedString)
				{
					int id = -1;
					if (Int32.TryParse(cellvalue, out id))
					{
						SharedStringItem item = GetItem(shareStringPart, id);
						if (item.Text != null)
						{
							//code to take the string value  
							cellvalue = item.Text.Text;
						}
						else if (item.InnerText != null)
						{
							cellvalue = item.InnerText;
						}
						else if (item.InnerXml != null)
						{
							cellvalue = item.InnerXml;
						}
					}
				}
			}
			return cellvalue;
		}

		/// <summary>
		/// 获取单元格的值
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="shareStringPartValues"></param>
		/// <returns></returns>
		public static string GetValue(this Cell cell, string[] shareStringPartValues)
		{
			if (cell == null)
				return null;
			string cellvalue = cell.InnerText;
			if (cell.DataType != null)
			{
				if (cell.DataType == CellValues.SharedString)
				{
					int id = -1;
					if (Int32.TryParse(cellvalue, out id))
					{
						cellvalue = shareStringPartValues[id];
					}
				}
			}
			return cellvalue;
		}

		/// <summary>
		/// 设置单元格值
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="value"></param>
		/// <param name="shareStringPart"></param>
		/// <param name="shareStringItemIndex"></param>
		/// <param name="styleIndex"></param>
		/// <returns></returns>
		public static Cell SetValue(this Cell cell, object value = null, SharedStringTablePart shareStringPart = null, int shareStringItemIndex = -1, uint styleIndex = 0)
		{
			if (value == null)
			{
				cell.CellValue = new CellValue();
				if (shareStringItemIndex != -1)
				{
					cell.CellValue = new CellValue(shareStringItemIndex.ToString());
					cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
				}
			}
			else if (value is string str)
			{
				if (shareStringPart == null)
				{
					cell.CellValue = new CellValue(str);
					cell.DataType = new EnumValue<CellValues>(CellValues.String);
				}
				else
				{
					// Insert the text into the SharedStringTablePart.
					int index = shareStringPart.GetOrInsertItem(str, false);
					// Set the value of cell
					cell.CellValue = new CellValue(index.ToString());
					cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
				}
			}
			else if (value is int || value is short || value is long ||
			  value is float || value is double || value is uint ||
			  value is ulong || value is ushort || value is decimal)
			{
				cell.CellValue = new CellValue(value.ToString());
				cell.DataType = new EnumValue<CellValues>(CellValues.Number);
			}
			else if (value is DateTime date)
			{
				cell.CellValue = new CellValue(date.ToString("yyyy-MM-dd HH:mm:ss")); // ISO 861
				cell.DataType = new EnumValue<CellValues>(CellValues.Date);
			}
			else if (value is XmlDocument xd)
			{
				if (shareStringPart == null)
				{
					throw new Exception("Param [shareStringPart] can't be null when value type is XmlDocument.");
				}
				else
				{
					int index = shareStringPart.GetOrInsertItem(xd.OuterXml, true);
					// Set the value of cell
					cell.CellValue = new CellValue(index.ToString());
					cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
				}
			}

			if (styleIndex != 0)
				cell.StyleIndex = styleIndex;

			return cell;
		}

		public static int GetOrInsertItem(this SharedStringTablePart shareStringPart, string content, bool isXml)
		{
			// If the part does not contain a SharedStringTable, create one.
			if (shareStringPart.SharedStringTable == null)
			{
				shareStringPart.SharedStringTable = new SharedStringTable();
			}

			int i = 0;

			foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
			{
				if ((!isXml && item.InnerText == content) || (isXml && item.OuterXml == content))
				{
					return i;
				}

				i++;
			}

			if (isXml)
				shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(content));
			else
				shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(content)));
			shareStringPart.SharedStringTable.Save();

			return i;
		}
		private static SharedStringItem GetItem(this SharedStringTablePart shareStringPart, int id)
		{
			return shareStringPart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
		}

		public static MergeCells GetOrCreateMergeCells(this Worksheet worksheet)
		{
			MergeCells mergeCells;
			if (worksheet.Elements<MergeCells>().Count() > 0)
			{
				mergeCells = worksheet.Elements<MergeCells>().First();
			}
			else
			{
				mergeCells = new MergeCells();

				// Insert a MergeCells object into the specified position.
				if (worksheet.Elements<CustomSheetView>().Count() > 0)
				{
					worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
				}
				else if (worksheet.Elements<DataConsolidate>().Count() > 0)
				{
					worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
				}
				else if (worksheet.Elements<SortState>().Count() > 0)
				{
					worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
				}
				else if (worksheet.Elements<AutoFilter>().Count() > 0)
				{
					worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
				}
				else if (worksheet.Elements<Scenarios>().Count() > 0)
				{
					worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
				}
				else if (worksheet.Elements<ProtectedRanges>().Count() > 0)
				{
					worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
				}
				else if (worksheet.Elements<SheetProtection>().Count() > 0)
				{
					worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
				}
				else if (worksheet.Elements<SheetCalculationProperties>().Count() > 0)
				{
					worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
				}
				else
				{
					worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
				}
				worksheet.Save();
			}
			return mergeCells;
		}

		public static void MergeTwoCells(this MergeCells mergeCells, string cell1Name, string cell2Name)
		{
			mergeCells.Append(new MergeCell() { Reference = new StringValue(cell1Name + ":" + cell2Name) });
		}

		public static IEnumerable<string> GetItemValues(this SharedStringTablePart shareStringPart)
		{
			foreach (var item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
			{
				if (item.Text != null)
				{
					//code to take the string value  
					yield return item.Text.Text;
				}
				else if (item.InnerText != null)
				{
					yield return item.InnerText;
				}
				else if (item.InnerXml != null)
				{
					yield return item.InnerXml;
				}
				else
				{
					yield return null;
				}
			};
		}

		public static XmlDocument GetCellAssociatedSharedStringItemXmlDocument(this SheetData sheetData, string columnName, uint rowIndex, SharedStringTablePart shareStringPart)
		{
			Cell cell = GetCell(sheetData, columnName, rowIndex);
			if (cell == null)
				return null;

			if (cell.DataType == CellValues.SharedString)
			{
				int id = -1;
				if (Int32.TryParse(cell.InnerText, out id))
				{
					SharedStringItem ssi = shareStringPart.GetItem(id);
					var doc = new XmlDocument();
					doc.LoadXml(ssi.OuterXml);
					return doc;
				}
			}
			return null;
		}
	}
}
