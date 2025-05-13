using System;
using System.Text;
using System.Collections.Generic;
using ExcelKit.Core.Constraint.Enums;

namespace ExcelKit.Core.ExcelRead
{
    /// <summary>
    /// 读取Sheet
    /// </summary>
    public class ReadSheetDicOptions
    {
        /// <summary>
        /// 表头数据所在行
        /// </summary>
        public uint HeadRow { get; set; } = 1;

        /// <summary>
        /// 数据起始行（从1开始计算）
        /// </summary>
        public uint DataStartRow { get; set; } = 2;

        /// <summary>
        /// 数据结束行（从1开始计算）
        /// </summary>
        public uint? DataEndRow { get; set; }

        /// 读取方式(默认按照Sheet索引读)
        /// </summary>
        public ReadWay ReadWay { get; set; } = ReadWay.SheetIndex;

        /// <summary>
        /// Sheet名称，默认Sheet1
        /// </summary>
        public string SheetName { get; set; } = "Sheet1";

        /// <summary>
        /// Sheet索引
        /// </summary>
        public ushort SheetIndex { get; set; } = 1;

        /// <summary>
        /// Excel中的列头信息（field:excel中列头字段 type:列的类型）
        /// </summary>
        /// <remarks>
        /// excel中某些单元格比如日期（OADate为一串数字）或者小数（有使用科学计数法的）读取出来数据不对，
        /// 所以此处增加了类型进来，用于转换为指定类型，这样看到的结果才是正确的
        /// </remarks>
        public (string field, ColumnType type, bool allowNull)[] ExcelFields { get; set; }

        /// <summary>
        /// 读取成功的行数据处理函数(Param1:读取出的数据  Param2：行索引)
        /// </summary>
        public Action<Dictionary<string, object>, uint> SucData { get; set; }

        /// <summary>
        /// 读取失败的行数据数据处理函数【失败的原始数据，失败数据信息汇总】
        /// </summary>
        public Action<Dictionary<string, object>, List<(string rowIndex, string columnName, string cellValue, string errorMsg)>> FailData { get; set; }

        /// <summary>
        /// 当直接传入Stream时，是否释放流(如果是指定的FilePath，则此选项不生效)
        /// 比如读取表头后，再根据表头读取具体内容，此时可指定为false不释放
        /// </summary>
        public bool IsDisposeStream { get; set; } = true;
    }
}
