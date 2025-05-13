using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using ExcelKit.Core.Attributes;
using ExcelKit.Core.Helpers;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using ExcelKit.Core.ExcelWrite;
using ExcelKit.Core.Infrastructure.Reflections;
using ExcelKit.Core.Constraint.Mappings;
using System.Threading;
using ExcelKit.Core.Constraint.Consts;
using ExcelKit.Core.Infrastructure.Exceptions;
using CellType = NPOI.SS.UserModel.CellType;

namespace ExcelKit.Core.Infrastructure
{
    /// <summary>
    /// 多阶段导出
    /// </summary>
    internal class MultiStageExporter
    {
        /// <summary>
        /// Sheet信号量
        /// </summary>
        static readonly Dictionary<string, SemaphoreSlim> _semaphores = new Dictionary<string, SemaphoreSlim>()
        {
            { SemaphoreKeyConst.GetSheet, new SemaphoreSlim(1) } ,{ SemaphoreKeyConst.CreateSheet, new SemaphoreSlim(1) }
        };

        /// <summary>
        /// workbook集(Key:fileid   Value:InnerSheetInfo)
        /// </summary>
        static ConcurrentDictionary<string, ExcelContextInfo> _workbooks = new ConcurrentDictionary<string, ExcelContextInfo>();

        /// <summary>
        /// 转换器接口默认名称
        /// </summary>
        static string ConverterInterfaceName = typeof(Converter.IExportConverter<>).Name.Split("`")[0];

        /// <summary>
        /// 去除SheetName中无效字符
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static string GetSafeSheetName(string sheetName)
        {
            Inspector.NotNullOrWhiteSpace(sheetName, "Sheet名称不能为空");
            return Regex.Replace(sheetName, "[\\[\\]\\^[\\]\\/\\-_*×――(^)|'$%~!@#$…&%￥—+=<>《》!！??？:：•`·、。，；,.;\"‘’“”-]", "_");
        }

        /// <summary>
        /// 获取Sheet
        /// </summary>
        /// <param name="fileId">文件标识</param>
        /// <param name="sheetName">Sheet名称</param>
        /// <returns></returns>
        //[System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.Synchronized)]
        static ISheet GetSheet(string fileId, string sheetName)
        {
            var workbook = _workbooks[fileId].Workbook;
            sheetName = GetSafeSheetName(sheetName);
            ISheet sheet = null;
            //SXSSFWorkbook获取SheetName方式有坑，最好用这种
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                var _sheet = workbook.GetSheetAt(i);
                if (_sheet.SheetName == sheetName)
                {
                    sheet = _sheet;
                    break;
                }
            }
            return sheet;

            //try
            //{
            //	_semaphores[SemaphoreKeyConst.GetSheet].Wait();

            //	ISheet sheet = null;
            //	//SXSSFWorkbook获取SheetName方式有坑，最好用这种
            //	for (int i = 0; i < workbook.NumberOfSheets; i++)
            //	{
            //		var _sheet = workbook.GetSheetAt(i);
            //		if (_sheet.SheetName == sheetName)
            //		{
            //			sheet = _sheet;
            //			break;
            //		}
            //	}
            //	return sheet;
            //}
            //finally
            //{
            //	_semaphores[SemaphoreKeyConst.GetSheet].Release();
            //}
        }

        /// <summary>
        /// 创建文件
        /// </summary>
        /// <param name="fileName">导出文件名称</param>
        /// <returns></returns>
        public static string CreateExcel(string fileName)
        {
            Inspector.NotNullOrWhiteSpace(fileName, "导出文件名称不能为空");

            var fileid = Guid.NewGuid().ToString("N");
            _workbooks.TryAdd(fileid, new ExcelContextInfo()
            {
                FileId = fileid,
                Workbook = null,
                BuildTime = DateTime.Now,
                ExportFileName = fileName
            });
            return fileid;
        }

        /// <summary>
        /// 创建Sheet
        /// </summary>
        /// <param name="fileId">文件标识</param>
        /// <param name="sheetName">Sheet名称</param>
        /// <param name="headers">表头行字段信息，key：表头列名称   value：列对应的Converter</param>
        //[System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.Synchronized)]
        public static ISheet CreteSheet(string fileId, string sheetName, List<ExcelKitAttribute> headers)
        {
            Inspector.Validation(!_workbooks.ContainsKey(fileId), "导出文件标识不存在");
            Inspector.Validation(headers == null, "Sheet表头信息不能为空");

            try
            {
                _semaphores[SemaphoreKeyConst.CreateSheet].Wait();

                headers = headers.OrderBy(t => t.Sort).ToList();

                //1.首次出现的Sheet记录下来（后面自动拆分Sheet要引用最初Sheet的信息）
                if (!_workbooks[fileId].SheetInfo.Exists(t => t.OriginSheetName == sheetName) && !sheetName.Contains(MultiStageExporterConst.INNER_SHEET_CHAR))
                {
                    _workbooks[fileId].SheetInfo.Add(new InnerSheetInfo()
                    {
                        SheetIndex = 0,
                        PropAttr = headers,
                        OriginSheetName = sheetName,
                        RealSheetName = sheetName,
                    });
                }
                else
                {
                    sheetName = sheetName.Replace(MultiStageExporterConst.INNER_SHEET_CHAR, "");

                    _workbooks[fileId].SheetInfo.Add(new InnerSheetInfo()
                    {
                        OriginSheetName = _workbooks[fileId].SheetInfo.First().OriginSheetName,
                        RealSheetName = sheetName,
                        PropAttr = headers,
                        SheetIndex = _workbooks[fileId].SheetInfo.Max(t => t.SheetIndex) + 1,
                        CellStyles = _workbooks[fileId].SheetInfo.First().CellStyles,
                    });
                }

                sheetName = GetSafeSheetName(sheetName);
                Inspector.NotNullOrWhiteSpace(sheetName, "Sheet名称为空或无效");

                //2.创建Workbook
                _workbooks[fileId].Workbook = _workbooks[fileId].Workbook ?? new NPOI.XSSF.Streaming.SXSSFWorkbook();

                //3.获取或创建Sheet
                var sheet = GetSheet(fileId, sheetName);
                if (sheet == null)
                {
                    var cellCount = 0;
                    sheet = _workbooks[fileId].Workbook.CreateSheet(sheetName);

                    var headerRow = sheet.CreateRow(0);

                    foreach (var head in headers)
                    {
                        //忽略的字段
                        if (head.IsIgnore)
                            continue;
                        //列宽设置
                        if (head.Width > 0)
                            sheet.SetColumnWidth(cellCount, head.Width * 256);
                        //首行冻结
                        if (head.HeadRowFrozen)
                            sheet.CreateFreezePane(headers.Count, 1);
                        //首行筛选
                        if (head.HeadRowFilter)
                            sheet.SetAutoFilter(new CellRangeAddress(0, 0, 0, headers.Count - 1));
                        //单元格样式
                        var cell = headerRow.CreateCell(cellCount++);
                        {
                            ICellStyle cellStyle = _workbooks[fileId].Workbook.CreateCellStyle();
                            cellStyle.Alignment = TextAlignMapping.MapAlign(head.Align);

                            //前景色
                            if (head.ForegroundColor != Constraint.Enums.DefineColor.None)
                            {
                                cellStyle.FillPattern = FillPattern.SolidForeground;
                                cellStyle.FillForegroundColor = ColorMapping.GetColorIndex(head.ForegroundColor);
                            }

                            //字体颜色
                            var font = _workbooks[fileId].Workbook.CreateFont();
                            font.FontHeightInPoints = 12;
                            font.FontName = "Calibri";
                            font.Color = ColorMapping.GetColorIndex(head.FontColor);
                            cellStyle.SetFont(font);

                            //记录单元格样式
                            if (_workbooks[fileId].SheetInfo.Exists(t => t.OriginSheetName == sheetName) && !sheetName.Contains(MultiStageExporterConst.INNER_SHEET_CHAR))
                            {
                                _workbooks[fileId].SheetInfo.FirstOrDefault(t => t.OriginSheetName == sheetName).CellStyles.Add(head.Code, cellStyle);
                            }

                            cell.CellStyle = cellStyle;//为单元格设置显示样式  
                        }
                        cell.SetCellValue(head.Desc);
                    }
                    _workbooks[fileId].SheetCount = _workbooks[fileId].Workbook.NumberOfSheets;
                }
                return sheet;
            }
            finally
            {
                _semaphores[SemaphoreKeyConst.CreateSheet].Release();
            }
        }

        /// <summary>
        /// 自动拆分Sheet
        /// </summary>
        private static void AutoSplitSheet(ref ISheet sheet, string fileId, string sheetName, uint autoSplit = MultiStageExporterConst.SheetMaxRowCount)
        {
            //获取传入的Sheet信息
            var lastSheetInfo = _workbooks[fileId].SheetInfo.OrderBy(t => t.SheetIndex).LastOrDefault(t => t.OriginSheetName == sheetName);
            //自动拆分Sheet时，由于传入的sheet名称是原始的sheet名称，故此处自动判断后定位到拆分后具体的sheet
            if (_workbooks[fileId].SheetInfo.Where(t => t.OriginSheetName == sheetName).Count() > 1 && lastSheetInfo.RealSheetName != sheetName)
            {
                sheet = GetSheet(fileId, InnerSheetInfo.BuildRealSheetName(lastSheetInfo.OriginSheetName, lastSheetInfo.SheetIndex));
            }

            //自动拆分Sheet
            if (lastSheetInfo.DataRowCount > autoSplit)
            {
                var lastSheet = GetSheet(fileId, InnerSheetInfo.BuildRealSheetName(lastSheetInfo.OriginSheetName, lastSheetInfo.SheetIndex));

                if (lastSheet != null && lastSheetInfo.DataRowCount <= autoSplit)
                {
                    sheet = lastSheet;
                }
                else
                {
                    var innerSheetName = InnerSheetInfo.BuildInnerSheetName(lastSheetInfo.OriginSheetName, lastSheetInfo.SheetIndex + 1);
                    sheet = CreteSheet(fileId, innerSheetName, lastSheetInfo.PropAttr);
                }
            }
        }

        /// <summary>
        /// 写入数据
        /// </summary>
        /// <param name="fileId">文件标识</param>
        /// <param name="sheetName">Sheet名称</param>
        /// <param name="data">数据</param>
        /// <param name="autoSplit">多少条数据拆分Sheet</param>
        /// <remarks>单Sheet默认最大1048575行，但直接使用这个会溢出，使用1048200不会</remarks>
        public static void AppendData<T>(string fileId, string sheetName, T data, uint autoSplit = MultiStageExporterConst.SheetMaxRowCount) where T : class, new()
        {
            Inspector.NotNull(data, "待导出的数据集不能为空");
            Inspector.Validation(!_workbooks.ContainsKey(fileId), $"导出文件标识不存在，请调用{nameof(CreateExcel)}创建");

            var sheet = GetSheet(fileId, sheetName) ?? throw new ExcelKitException($"导出文件中不存在名为 {sheetName} 的Sheet，请调用{nameof(CreteSheet)}创建");
            var originSheetName = sheetName;

            //自动拆分Sheet
            autoSplit = autoSplit > MultiStageExporterConst.SheetMaxRowCount ? MultiStageExporterConst.SheetMaxRowCount : autoSplit;
            AutoSplitSheet(ref sheet, fileId, sheetName, autoSplit);

            //当前Sheet信息(注意此处要用真实Sheet名称获取Sheet信息 ，ref sheet中已返回实际正在操作的Sheet)
            var curSheetInfo = _workbooks[fileId].SheetInfo
                .Where(t => t.OriginSheetName == sheetName)
                .OrderBy(t => t.SheetIndex)
                .LastOrDefault(t => t.RealSheetName == sheet.SheetName);

            //写入数据
            var cellCount = 0;
            curSheetInfo.IncrementDataRowCount();
            var dataRow = sheet.CreateRow(curSheetInfo.DataRowCount - 1);
            var sortedProps = ReflectionHelper.NewInstance.GetSortedExportProps<T>();
            foreach (var item in sortedProps)
            {
                object value = null;
                if (item.attr.Converter == null)
                {
                    value = item.prop.GetValue(data) ?? "";
                }
                else
                {
                    try
                    {
                        var convertInterfaces = item.attr.Converter.GetInterfaces();
                        var isConverter = Array.Exists(convertInterfaces, t => t.GetGenericTypeDefinition().Name.Contains(ConverterInterfaceName));
                        if (isConverter)
                        {
                            var overrideMethodTypes = new List<Type>() { item.prop.PropertyType };
                            var parameters = new List<object>() { NumToDecimalMapping.IfNeedToDecimal(item.prop.GetValue(data)) };
                            //提取ConverterParam上的参数给Convert（默认的接口是一个泛型参数，不是默认则说明需要指定ConverterParam）
                            if (convertInterfaces.Count(t => t.GetGenericTypeDefinition() != typeof(Converter.IExportConverter<>)) > 0)
                            {
                                parameters.Add(item.attr.ConverterParam);
                                overrideMethodTypes.Add(item.attr.ConverterParam == null ? typeof(string) : item.attr.ConverterParam.GetType());
                            }
                            value = item.attr.Converter.GetMethod("Convert", overrideMethodTypes.ToArray()).Invoke(Activator.CreateInstance(item.attr.Converter), parameters.ToArray());
                        }
                        else
                        {
                            value = $"Excel列 {item.attr.Desc} 字段上指定的Converter类型不是有效的Converter";
                        }
                    }
                    catch (Exception ex)
                    {
                        value = $"Convert Is 【{item.attr.Converter.AssemblyQualifiedName}】，Value Is {item.prop.GetValue(data)?.ToString()}，Convert Fail，ExceptionInfo:{ex.Message}";
                    }
                }

                value = value is DateTime ? ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss") : value;
                var cell = dataRow.CreateCell(cellCount++);
                {
                    var thisSheet = _workbooks[fileId].SheetInfo.FirstOrDefault(t => t.OriginSheetName == originSheetName);
                    cell.CellStyle = thisSheet.CellStyles[item.attr.Code];//为单元格设置显示样式  
                }
                if (value is byte || value is int || value is long || value is short || value is float || value is double || value is decimal)
                {
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(Convert.ToDouble(value));
                }
                else
                {
                    cell.SetCellValue(value?.ToString());
                }
            }
        }

        /// <summary>
        /// 写入数据
        /// </summary>
        /// <param name="fileId">文件标识</param>
        /// <param name="sheetName">Sheet名称</param>
        /// <param name="rowData">单行数据</param>
        /// <param name="autoSplit">多少条数据拆分Sheet</param>
        /// <remarks>单Sheet默认最大1048575行，但直接使用这个会溢出，使用1048200不会</remarks>
        public static void AppendData(string fileId, string sheetName, Dictionary<string, object> rowData, uint autoSplit = MultiStageExporterConst.SheetMaxRowCount)
        {
            Inspector.NotNull(rowData, "待导出的数据集不能为空");
            Inspector.Validation(!_workbooks.ContainsKey(fileId), $"导出文件标识不存在，请调用{nameof(CreateExcel)}创建");

            var sheet = GetSheet(fileId, sheetName) ?? throw new ExcelKitException($"导出文件中不存在名为 {sheetName} 的Sheet，请调用{nameof(CreteSheet)}创建");
            var originSheetName = sheetName;

            //自动拆分Sheet
            autoSplit = autoSplit > MultiStageExporterConst.SheetMaxRowCount ? MultiStageExporterConst.SheetMaxRowCount : autoSplit;
            AutoSplitSheet(ref sheet, fileId, sheetName, autoSplit);

            //当前Sheet信息(注意此处要用真实Sheet名称获取Sheet信息 ，ref sheet中已返回实际正在操作的Sheet)
            var curSheetInfo = _workbooks[fileId].SheetInfo
                .Where(t => t.OriginSheetName == sheetName)
                .OrderBy(t => t.SheetIndex)
                .LastOrDefault(t => t.RealSheetName == sheet.SheetName);

            //写入数据
            var cellCount = 0;
            curSheetInfo.IncrementDataRowCount();
            var dataRow = sheet.CreateRow(curSheetInfo.DataRowCount - 1);
            var headers = _workbooks[fileId].SheetInfo.FirstOrDefault(t => t.OriginSheetName == sheetName).PropAttr;

            var index = 0;
            foreach (var item in headers)
            {
                object value = null;
                var attributeElement = headers.ElementAt(index);

                if (attributeElement.Converter == null)
                {
                    value = rowData.ContainsKey(item.Code) ? rowData[item.Code] : "";
                }
                else
                {
                    try
                    {
                        var convertInterfaces = attributeElement.Converter.GetInterfaces();
                        var isConverter = Array.Exists(convertInterfaces, t => t.GetGenericTypeDefinition().Name.Contains(ConverterInterfaceName));
                        if (isConverter)
                        {
                            var parameters = new List<object>() { NumToDecimalMapping.IfNeedToDecimal(rowData[item.Code]) };
                            //提取ConverterParam上的参数给Convert（默认的接口是一个泛型参数，不是默认则说明需要指定ConverterParam）
                            if (convertInterfaces.Count(t => t.GetGenericTypeDefinition() != typeof(Converter.IExportConverter<>)) > 0)
                            {
                                //TODO：此处如果ConverterParam指定的参数错误导致拆分后参数数量不对，会导致反射调用失败，最好是自动补齐参数对应类型的数量
                                //parameters.AddRange(attributeElement.ConverterParam.ToString().Split("|")?.ToList<object>());
                                parameters.Add(attributeElement.ConverterParam);
                            }
                            value = attributeElement.Converter.GetMethod("Convert", new Type[] { rowData[item.Code].GetType(), attributeElement.ConverterParam == null ? typeof(string) : attributeElement.ConverterParam.GetType() }).Invoke(Activator.CreateInstance(attributeElement.Converter), parameters.ToArray());
                        }
                        else
                        {
                            value = $"Excel列 {attributeElement.Desc} 字段上指定的Converter类型不是有效的Converter";
                        }
                    }
                    catch (Exception ex)
                    {
                        value = $"Convert Is 【{attributeElement.Converter.AssemblyQualifiedName}】，Value Is {rowData[item.Code]?.ToString()}，Convert Fail，ExceptionInfo:{ex.Message}";
                    }
                }

                value = value is DateTime ? ((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss") : value;
                var cell = dataRow.CreateCell(cellCount++);
                {
                    var thisSheet = _workbooks[fileId].SheetInfo.FirstOrDefault(t => t.OriginSheetName == originSheetName);
                    cell.CellStyle = thisSheet.CellStyles[attributeElement.Code];//为单元格设置显示样式  
                }
                if (value is byte || value is int || value is long || value is short || value is float || value is double || value is decimal)
                {
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(Convert.ToDouble(value));
                }
                else
                {
                    cell.SetCellValue(value?.ToString());
                }
                index++;
            }
        }

        /// <summary>
        /// 生成并保存Excel
        /// </summary>
        /// <param name="fileId">文件标识</param>
        /// <param name="saveForder">保存路径</param>
        /// <returns>文件路径</returns>
        public static string Save(string fileId, string saveForder = null)
        {
            try
            {
                var filePath = string.IsNullOrWhiteSpace(saveForder)
                               ? Path.Combine(AppContext.BaseDirectory, $"{_workbooks[fileId].ExportFileName.Replace(".xlsx", "")}.xlsx")
                               : Path.Combine(saveForder, $"{_workbooks[fileId].ExportFileName.Replace(".xlsx", "")}.xlsx");

                using (FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite))
                {
                    _workbooks[fileId].Workbook.Write(stream);
                }

                return filePath;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                _workbooks.Remove(fileId, out ExcelContextInfo excelInfo);
                excelInfo.Workbook.Close();
            }
        }

        /// <summary>
        /// 生成Excel流
        /// </summary>
        /// <param name="fileId">文件标识</param>
        public static OutExcelInfo Generate(string fileId)
        {
            try
            {
                byte[] buffer;
                using (var ms = new MemoryStream())
                {
                    _workbooks[fileId].Workbook.Write(ms);
                    buffer = ms.ToArray();
                }
                Stream stream = new MemoryStream(buffer);

                return new OutExcelInfo()
                {
                    Stream = stream,
                    FileName = $"{_workbooks[fileId].ExportFileName.Replace(".xlsx", "")}.xlsx"
                };
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                _workbooks.Remove(fileId, out ExcelContextInfo excelInfo);
                excelInfo.Workbook.Close();
            }
        }

        /// <summary>
        /// 清理资源
        /// </summary>
        /// <param name="fileId">文件标识</param>
        public static void Dispose(string fileId)
        {
            _workbooks.TryRemove(fileId, out ExcelContextInfo excelInfo);
            excelInfo?.Workbook?.Close();
        }
    }
}
