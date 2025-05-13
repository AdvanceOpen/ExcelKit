using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using ExcelKit.Core.Helpers;
using ExcelKit.Core.Extensions;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using ExcelKit.Core.Infrastructure.Exceptions;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using ExcelKit.Core.Constraint.Enums;
using ExcelKit.Core.Attributes;
using System.Reflection;
using ExcelKit.Core.Infrastructure.Reflections;
using ExcelKit.Core.Constraint.Mappings;

namespace ExcelKit.Core.ExcelRead
{
    internal class ReadExcelContext : IReadExcelContext
    {
        //默认的Excel列头，ReadRows系列方法使用；此组件在遇到空单元格时会获取不到对应单元格直接返回有值的单元格，
        //此处定义默认的，用于数据补齐，不然ReadRows系列方法会数据错位（每行空单元格位置不一样，数组中单行数据位置就不一样）
        static readonly string[] _excelColumn = new string[]
        {
            "A", "B", "C", "D", "E", "F", "G", "H", "I","J", "K", "L", "M", "N", "O", "P", "Q", "R","S","T","U","V","W","X","Y","Z",
            "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI","AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR","AS","AT","AU","AV","AW","AX","AY","AZ",
        };

        private void CheckExcelInfo(string filePath)
        {
            Inspector.NotNull(filePath, "Excel文件不能为空");
            Inspector.Validation(!Path.GetFileName(filePath).EndsWith(".xlsx"), "Excel文件类型错误，当前仅支持xlsx格式");
        }

        public IEnumerable<string> GetSheetNames(string filePath)
        {
            this.CheckExcelInfo(filePath);

            FileStream fileStream = null;
            SpreadsheetDocument sheetDoc = null;
            try
            {
                fileStream = File.OpenRead(filePath);
                sheetDoc = SpreadsheetDocument.Open(fileStream, false);
                return sheetDoc.WorkbookPart.GetSheetNames();
            }
            finally
            {
                sheetDoc?.Dispose();
                fileStream?.Dispose();
            }
        }

        public IEnumerable<string> GetSheetNames(Stream stream, bool disposeStream = true)
        {
            Inspector.NotNull(stream, "Excel文件流不能为空");

            Func<SpreadsheetDocument, IEnumerable<string>> readSheetNames = (sheetDoc) =>
            {
                WorkbookPart workbookPart = sheetDoc.WorkbookPart;
                return workbookPart.GetSheetNames();
            };

            if (disposeStream)
            {
                using (var sheetDoc = SpreadsheetDocument.Open(stream, false))
                {
                    return readSheetNames(sheetDoc);
                }
            }
            else
            {
                return readSheetNames(SpreadsheetDocument.Open(stream, false));
            }
        }

        public void ReadRows(string filePath, ReadRowsOptions option)
        {
            this.CheckExcelInfo(filePath);
            using (var stream = File.OpenRead(filePath))
            {
                this.ReadRows(stream, option);
            }
        }

        public void ReadRows(Stream stream, ReadRowsOptions option)
        {
            Inspector.NotNull(stream, "Excel文件流不能为空");
            Inspector.NotNull(option, $"{nameof(ReadRowsOptions)}不能为null");
            if (option.ReadWay == ReadWay.SheetIndex)
                Inspector.MoreThanOrEqual(option.SheetIndex, 1, "Sheet索引至少从1开始");
            if (option.ReadWay == ReadWay.SheetName)
                Inspector.NotNullOrWhiteSpace(option.SheetName, "Sheet名称不能为空");

            //匹配SheetName
            var sheetName = "";
            {
                var sheetNames = this.GetSheetNames(stream, false);
                Inspector.Validation(option.ReadWay == ReadWay.SheetIndex && option.SheetIndex > sheetNames.Count(), $"指定的Sheet索引 {option.SheetIndex} 无效，实际只存在{sheetNames.Count()}个Sheet");
                Inspector.Validation(option.ReadWay == ReadWay.SheetName && !sheetNames.Contains(option.SheetName), $"指定的Sheet名称 {option.SheetName} 不存在");
                sheetName = option.ReadWay switch { ReadWay.SheetIndex => sheetNames.ElementAt(option.SheetIndex - 1), ReadWay.SheetName => option.SheetName };
            }

            //包含返回空单元格校验
            Inspector.Validation(option.ReadEmptyCell && (option.ColumnHeaders == null || option.ColumnHeaders.Count() == 0), $"读取要包含空格列时，必须指定读取的列头信息，{nameof(option.ColumnHeaders)}不能为空");

            if (option.ReadEmptyCell)
            {
                foreach (var cellRef in option.ColumnHeaders)
                {
                    Inspector.Validation(!_excelColumn.Contains(cellRef), $"{nameof(option.ColumnHeaders)}中 {cellRef} Excel列头无效");
                }
            }

            using (var sheetDoc = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbookPart = sheetDoc.WorkbookPart;
                SharedStringTablePart shareStringPart;
                if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                else
                    shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();

                string[] shareStringItemValues = shareStringPart.GetItemValues().ToArray();

                //目标Sheet的Rid
                string rId = workbookPart.Workbook.Sheets?.Cast<Sheet>()?.FirstOrDefault(t => t.Name.Value == sheetName)?.Id?.Value;
                Inspector.NotNullOrWhiteSpace(rId, $"不存在名为：{sheetName} 的Sheet");

                //读取多个Sheet
                foreach (var worksheetPart in workbookPart.WorksheetParts?.Reverse())
                {
                    //是否是指定Sheet的Rid
                    string partRelationshipId = workbookPart.GetIdOfPart(worksheetPart);
                    if (partRelationshipId != rId) continue;

                    OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
                    //数据行
                    var list = new List<(string cellRef, string value)>();

                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(Worksheet))
                        {
                            reader.ReadFirstChild();
                        }

                        if (reader.ElementType == typeof(Row))
                        {
                            var row = (Row)reader.LoadCurrentElement();
                            if (row.RowIndex < option.DataStartRow)
                                continue;
                            if (option.DataEndRow.HasValue && row.RowIndex > option.DataEndRow)
                                break;

                            list.Clear();

                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                if (cell.CellReference == null || !cell.CellReference.HasValue)
                                    continue;

                                //无数字的列头名（A1、B2、C1中的A、B、C）
                                var loopCellRef = StringHelper.RemoveNumber(cell.CellReference);
                                if (option.ReadEmptyCell && !option.ColumnHeaders.Contains(loopCellRef))
                                    continue;

                                string value = cell.GetValue(shareStringItemValues);
                                list.Add((loopCellRef, value));
                            }

                            //包含空单元格，实际却没有则补齐数据
                            if (option.ReadEmptyCell)
                            {
                                var willAdd = option.ColumnHeaders.Except(list.Select(t => t.cellRef));
                                list.AddRange(willAdd.Select(item => (item, "")));
                                list = list.OrderBy(t => t.cellRef).ToList();
                            }

                            option.RowData?.Invoke(list.Select(t => t.value).ToList());
                        }
                    }
                }
            }
            if (option.IsDisposeStream) stream.Dispose();
        }

        public void ReadSheet(string filePath, ReadSheetDicOptions option)
        {
            this.CheckExcelInfo(filePath);
            using (var stream = File.OpenRead(filePath))
            {
                this.ReadSheet(stream, option);
            }
        }

        public void ReadSheet(Stream stream, ReadSheetDicOptions option)
        {
            Inspector.NotNull(stream, "Excel文件流不能为空");
            Inspector.NotNull(option, $"{nameof(ReadSheetDicOptions)} can not be null");
            Inspector.Validation(option.ExcelFields == null || option.ExcelFields.Length == 0 || option.ExcelFields.Count(t => string.IsNullOrWhiteSpace(t.field)) > 0, "Excel中的列头信息不能为空或存在为空的列名");

            //匹配SheetName
            var sheetName = "";
            {
                var sheetNames = this.GetSheetNames(stream, false);
                Inspector.Validation(option.ReadWay == ReadWay.SheetIndex && option.SheetIndex > sheetNames.Count(), $"指定的SheetIndex {option.SheetIndex} 无效，实际只存在{sheetNames.Count()}个Sheet");
                Inspector.Validation(option.ReadWay == ReadWay.SheetName && !sheetNames.Contains(option.SheetName), $"指定的SheetName {option.SheetName} 不存在");
                sheetName = option.ReadWay switch { ReadWay.SheetIndex => sheetNames.ElementAt(option.SheetIndex - 1), ReadWay.SheetName => option.SheetName };
            }

            //Excel中的表头列信息(index：集合中元素的位置，cellRef：单元格的A1 B1中的A  B这种)
            var fieldLoc = new List<(int index, string excelField, ColumnType columnType, bool allowNull, string cellRef)>();
            {
                for (int index = 0; index < option.ExcelFields.Count(); index++)
                {
                    Inspector.Validation(fieldLoc.Exists(t => t.excelField == option.ExcelFields[index].field?.Trim()), "指定读取的 ExcelFields 中存在相同的列名");
                    fieldLoc.Add((index, option.ExcelFields[index].field, option.ExcelFields[index].type, option.ExcelFields[index].allowNull, ""));
                }
            }

            using (var sheetDoc = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbookPart = sheetDoc.WorkbookPart;
                //1.目标Sheet的Rid是否存在
                string rId = workbookPart.Workbook.Sheets?.Cast<Sheet>()?.FirstOrDefault(t => t.Name.Value == sheetName)?.Id?.Value;
                Inspector.NotNullOrWhiteSpace(rId, $"不存在名为 {sheetName} 的Sheet");

                SharedStringTablePart shareStringPart;
                if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                else
                    shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();

                string[] shareStringItemValues = shareStringPart.GetItemValues().ToArray();

                //2.反转Sheet顺序
                foreach (var workSheetPart in workbookPart.WorksheetParts?.Reverse())
                {
                    //是否是指定Sheet的Rid，不是则忽略
                    string partRelationshipId = workbookPart.GetIdOfPart(workSheetPart);
                    if (partRelationshipId != rId) continue;

                    //读取失败的原始数据信息
                    (Dictionary<string, object> odata, List<(string rowIndex, string columnName, string cellValue, string errorMsg)> failInfos) failRowData =
                        (new Dictionary<string, object>(), new List<(string rowIndex, string columnName, string cellValue, string errorMsg)>());

                    //创建Reader
                    OpenXmlReader reader = OpenXmlReader.Create(workSheetPart);
                    //工具类实例
                    var reflection = ReflectionHelper.NewInstance;

                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(Worksheet))
                        {
                            reader.ReadFirstChild();
                        }

                        if (reader.ElementType == typeof(Row))
                        {
                            var row = (Row)reader.LoadCurrentElement();

                            //3.读取表头列，匹配字段信息
                            if (row.RowIndex == option.HeadRow)
                            {
                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    if (cell.CellReference != null && cell.CellReference.HasValue)
                                    {
                                        //excel中的表头列字段
                                        string excelField = cell.GetValue(shareStringItemValues);
                                        if (fieldLoc.Exists(t => t.excelField == excelField))
                                        {
                                            var fieldInfo = fieldLoc.FirstOrDefault(t => t.excelField == excelField);
                                            fieldInfo.cellRef = StringHelper.RemoveNumber(cell.CellReference);
                                            fieldLoc[fieldInfo.index] = fieldInfo;
                                        }
                                    }
                                }
                                //实体上定义了ExcelKit特性的字段未在Excel中匹配到
                                var unMatchedField = fieldLoc.Where(t => string.IsNullOrWhiteSpace(t.cellRef));
                                if (unMatchedField.Count() > 0)
                                {
                                    var unmatchFields = string.Join("、", unMatchedField.Select(t => t.excelField));
                                    var msg = $"指定的ExcelFields中的字段【{unmatchFields}】不存在于Excel中";
                                    throw new ExcelKitException(msg);
                                }
                                continue;
                            }

                            if (row.RowIndex < option.DataStartRow)
                                continue;
                            if (option.DataEndRow.HasValue && row.RowIndex > option.DataEndRow)
                                break;

                            //读取到的每行数据
                            var rowData = new Dictionary<string, object>();
                            //excel原始数据
                            failRowData.odata.Clear();
                            //失败信息
                            failRowData.failInfos.Clear();
                            //是否读取成功
                            var readSuc = true;

                            //4. row.Elements<Cell>()获取出来的会自动跳过为空的单元格
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                //4.1 跳过cell引用为空的
                                if (cell.CellReference == null || !cell.CellReference.HasValue)
                                {
                                    continue;
                                }

                                //4.2 当前循环的cell列位置(不含数字)
                                var loopCellRef = StringHelper.RemoveNumber(cell.CellReference);
                                //不存在或匹配列信息不一致的跳过
                                var fieldInfo = fieldLoc.FirstOrDefault(t => t.cellRef.Equals(loopCellRef, StringComparison.OrdinalIgnoreCase));
                                if (fieldInfo == (0, null, 0, false, null) || !loopCellRef.Equals(fieldInfo.cellRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }

                                //Excel中读取到的值
                                string readVal = null;
                                try
                                {
                                    readVal = cell.GetValue(shareStringItemValues);
                                    Inspector.Validation(!fieldInfo.allowNull && string.IsNullOrWhiteSpace(readVal), $"Excel中列 {fieldInfo.excelField} 为必填项");

                                    object value = ColumnTypeMapping.Convert(readVal, fieldInfo.columnType, fieldInfo.allowNull);
                                    rowData.Add(fieldInfo.excelField, value);
                                }
                                catch (Exception ex)
                                {
                                    readSuc = false;
                                    failRowData.failInfos.Add((row.RowIndex, fieldInfo.excelField, readVal?.ToString(), ex.Message));
                                }
                            }

                            //5.单元格为空缺失的key补齐（这样做key的顺序和原始的不一致了，有需求时可以使用header上面的cellRef排序解决，为了读取速度此处暂不做）
                            var lackKeys = fieldLoc.Select(t => t.excelField).Except(rowData.Keys);
                            foreach (var lackKey in lackKeys)
                            {
                                rowData.TryAdd(lackKey, null);
                            }

                            //读取成功执行
                            if (readSuc)
                                option.SucData?.Invoke(rowData, row.RowIndex.Value);
                            else
                                option.FailData?.Invoke(failRowData.odata, failRowData.failInfos);
                        }
                    }
                }
            }
            if (option.IsDisposeStream) stream.Dispose();
        }

        public void ReadSheet<T>(string filePath, ReadSheetOptions<T> option) where T : class, new()
        {
            this.CheckExcelInfo(filePath);
            using (var stream = File.OpenRead(filePath))
            {
                this.ReadSheet<T>(stream, option);
            }
        }

        public void ReadSheet<T>(Stream stream, ReadSheetOptions<T> option) where T : class, new()
        {
            Inspector.NotNull(stream, "Excel文件流不能为空");
            Inspector.NotNull(option, $"{nameof(ReadSheetOptions<T>)} can not be null");

            //匹配SheetName
            var sheetName = "";
            {
                var sheetNames = this.GetSheetNames(stream, false);
                Inspector.Validation(option.ReadWay == ReadWay.SheetIndex && option.SheetIndex > sheetNames.Count(), $"指定的SheetIndex {option.SheetIndex} 无效，实际只存在{sheetNames.Count()}个Sheet");
                Inspector.Validation(option.ReadWay == ReadWay.SheetName && !sheetNames.Contains(option.SheetName), $"指定的SheetName {option.SheetName} 不存在");
                sheetName = option.ReadWay switch { ReadWay.SheetIndex => sheetNames.ElementAt(option.SheetIndex - 1), ReadWay.SheetName => option.SheetName };
            }

            //Excel中的表头列与实体类中字段映射(index：集合中元素的位置，cellRef：单元格的A1 B1中的A  B这种)
            var fieldLoc = new List<(int index, string excelField, string cellRef, string classField, bool allowNull, PropertyInfo prop)>();
            {
                var props = ReflectionHelper.NewInstance.GetSortedReadProps<T>().Select(t => t.prop).ToList();
                for (int index = 0; index < props.Count(); index++)
                {
                    var attribute = props[index].GetCustomAttribute<ExcelKitAttribute>();
                    fieldLoc.Add((index, attribute.Desc, "", attribute.Code ?? props[index].Name, attribute.AllowNull, props[index]));
                }
            }

            using (var sheetDoc = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbookPart = sheetDoc.WorkbookPart;
                //1.目标Sheet的Rid是否存在
                string rId = workbookPart.Workbook.Sheets?.Cast<Sheet>()?.FirstOrDefault(t => t.Name.Value == sheetName)?.Id?.Value;
                Inspector.NotNullOrWhiteSpace(rId, $"不存在名为：{sheetName} 的Sheet");

                SharedStringTablePart shareStringPart;
                if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                else
                    shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();

                string[] shareStringItemValues = shareStringPart.GetItemValues().ToArray();

                //2.反转Sheet顺序
                foreach (var workSheetPart in workbookPart.WorksheetParts?.Reverse())
                {
                    //是否是指定Sheet的Rid
                    string partRelationshipId = workbookPart.GetIdOfPart(workSheetPart);
                    if (partRelationshipId != rId) continue;

                    //读取失败的原始数据信息
                    (Dictionary<string, object> odata, List<(string rowIndex, string columnName, string cellValue, string errorMsg)> failInfos) failRowData =
                        (new Dictionary<string, object>(), new List<(string rowIndex, string columnName, string cellValue, string errorMsg)>());

                    //创建Reader
                    OpenXmlReader reader = OpenXmlReader.Create(workSheetPart);
                    //工具类实例
                    var reflection = ReflectionHelper.NewInstance;

                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(Worksheet))
                        {
                            reader.ReadFirstChild();
                        }

                        if (reader.ElementType == typeof(Row))
                        {
                            var row = (Row)reader.LoadCurrentElement();

                            //3.读取表头列，匹配字段信息
                            if (row.RowIndex == option.HeadRow)
                            {
                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    if (cell.CellReference != null && cell.CellReference.HasValue)
                                    {
                                        //excel中的表头列字段
                                        string outerCode = cell.GetValue(shareStringItemValues);
                                        if (fieldLoc.Exists(t => t.excelField == outerCode))
                                        {
                                            var fieldInfo = fieldLoc.FirstOrDefault(t => t.excelField == outerCode);
                                            fieldInfo.cellRef = StringHelper.RemoveNumber(cell.CellReference);
                                            fieldLoc[fieldInfo.index] = fieldInfo;
                                        }
                                    }
                                }
                                //实体上定义了ExcelKit特性的字段未在Excel中匹配到
                                var unMatchedField = fieldLoc.Where(t => string.IsNullOrWhiteSpace(t.cellRef));
                                if (unMatchedField.Count() > 0)
                                {
                                    var obj = unMatchedField.FirstOrDefault();
                                    var msg = $"{typeof(T).Name}中的字段{obj.classField}特性上指定的Desc：{obj.excelField}  未在Excel列头中匹配到";
                                    throw new ExcelKitException(msg);
                                }
                                continue;
                            }

                            if (row.RowIndex < option.DataStartRow)
                                continue;
                            if (option.DataEndRow.HasValue && row.RowIndex > option.DataEndRow)
                                break;

                            //读取到的每行数据
                            T model = new T();
                            //excel原始数据
                            failRowData.odata.Clear();
                            //失败信息
                            failRowData.failInfos.Clear();
                            //是否读取成功
                            var readSuc = true;

                            //4. row.Elements<Cell>()获取出来的会自动跳过为空的单元格
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                //4.1 跳过cell引用为空的
                                if (cell.CellReference == null || !cell.CellReference.HasValue)
                                {
                                    continue;
                                }

                                //4.2 当前循环的cell列位置(不含数字)
                                var loopCellRef = StringHelper.RemoveNumber(cell.CellReference);
                                //不存在或匹配列信息不一致的跳过
                                var fieldInfo = fieldLoc.FirstOrDefault(t => t.cellRef.Equals(loopCellRef, StringComparison.OrdinalIgnoreCase));
                                if (fieldInfo == (0, null, null, null, false, null) || !loopCellRef.Equals(fieldInfo.cellRef, StringComparison.OrdinalIgnoreCase))
                                {
                                    continue;
                                }

                                //Excel中读取到的值
                                string value = cell.GetValue(shareStringItemValues);
                                Inspector.Validation(!fieldInfo.allowNull && string.IsNullOrWhiteSpace(value), $"Excel中列 {fieldInfo.excelField} 为必填项");

                                try
                                {
                                    failRowData.odata.Add(fieldInfo.excelField, value);
                                    reflection.SetValue(fieldInfo.prop, ref model, value, fieldInfo.allowNull);
                                }
                                catch (Exception ex)
                                {
                                    readSuc = false;
                                    failRowData.failInfos.Add((row.RowIndex, fieldInfo.excelField, value?.ToString(), ex.Message));
                                }
                            }

                            //读取成功执行
                            if (readSuc)
                                option.SucData?.Invoke(model, row.RowIndex.Value);
                            else
                                option.FailData?.Invoke(failRowData.odata, failRowData.failInfos);
                        }
                    }
                }
            }
            if (option.IsDisposeStream) stream.Dispose();
        }

        public int ReadSheetRowsCount(string filePath, ReadSheetRowsCountOptions option)
        {
            this.CheckExcelInfo(filePath);
            using (var stream = File.OpenRead(filePath))
            {
                return this.ReadSheetRowsCount(stream, option);
            }
        }

        public int ReadSheetRowsCount(Stream stream, ReadSheetRowsCountOptions option)
        {
            Inspector.NotNull(stream, "Excel文件流不能为空");
            Inspector.NotNull(option, $"{nameof(ReadSheetRowsCountOptions)} can not be null");

            //匹配SheetName
            var sheetName = "";
            {
                var sheetNames = this.GetSheetNames(stream, false);
                Inspector.Validation(option.ReadWay == ReadWay.SheetIndex && option.SheetIndex > sheetNames.Count(), $"指定的SheetIndex {option.SheetIndex} 无效，实际只存在{sheetNames.Count()}个Sheet");
                Inspector.Validation(option.ReadWay == ReadWay.SheetName && !sheetNames.Contains(option.SheetName), $"指定的SheetName {option.SheetName} 不存在");
                sheetName = option.ReadWay switch { ReadWay.SheetIndex => sheetNames.ElementAt(option.SheetIndex - 1), ReadWay.SheetName => option.SheetName };
            }

            int rowsCount = 0;

            using (var sheetDoc = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbookPart = sheetDoc.WorkbookPart;
                //1.目标Sheet的Rid是否存在
                string rId = workbookPart.Workbook.Sheets?.Cast<Sheet>()?.FirstOrDefault(t => t.Name.Value == sheetName)?.Id?.Value;
                Inspector.NotNullOrWhiteSpace(rId, $"不存在名为：{sheetName} 的Sheet");

                SharedStringTablePart shareStringPart;
                if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                else
                    shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();

                string[] shareStringItemValues = shareStringPart.GetItemValues().ToArray();

                //2.反转Sheet顺序
                foreach (var workSheetPart in workbookPart.WorksheetParts?.Reverse())
                {
                    //是否是指定Sheet的Rid
                    string partRelationshipId = workbookPart.GetIdOfPart(workSheetPart);
                    if (partRelationshipId != rId) continue;

                    //创建Reader
                    OpenXmlReader reader = OpenXmlReader.Create(workSheetPart);

                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(Worksheet))
                        {
                            reader.ReadFirstChild();
                        }

                        if (reader.ElementType == typeof(Row))
                        {
                            var row = (Row)reader.LoadCurrentElement();
                            if (option.ContainsEmptyRow)
                            {
                                rowsCount++;
                                continue;
                            }

                            var hasNotEmptyCell = false;
                            //4. row.Elements<Cell>()获取出来的会自动跳过为空的单元格
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                //4.1 跳过cell引用为空的
                                if (cell.CellReference == null || !cell.CellReference.HasValue)
                                    continue;

                                //Excel中读取到的值
                                string value = cell.GetValue(shareStringItemValues);
                                if (!string.IsNullOrWhiteSpace(value))
                                {
                                    hasNotEmptyCell = true;
                                    break;
                                }
                            }
                            if (hasNotEmptyCell) rowsCount++;
                        }
                    }
                }
            }
            if (option.IsDisposeStream) stream.Dispose();
            return rowsCount;
        }
    }
}
