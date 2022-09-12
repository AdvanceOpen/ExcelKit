using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using ExcelKit.Core.Attributes;
using ExcelKit.Core.Constraint.Consts;
using ExcelKit.Core.Helpers;
using NPOI.SS.UserModel;

namespace ExcelKit.Core.ExcelWrite
{
    /// <summary>
    /// Sheet信息
    /// </summary>
    internal class InnerSheetInfo
    {
        /// <summary>
        /// 同一个Sheet名称自动拆分出的所有Sheet的TypeName相同
        /// </summary>
        public string OriginSheetName { get; set; }

        /// <summary>
        /// 真实的Sheet名称
        /// </summary>
        public string RealSheetName { get; set; }

        /// <summary>
        /// 字段信息
        /// </summary>
        public List<ExcelKitAttribute> PropAttr { get; set; }

        /// <summary>z
        /// 单元格样式
        /// </summary>
        public Dictionary<string, ICellStyle> CellStyles = new Dictionary<string, ICellStyle>();

        /// <summary>
        /// 该Sheet类型下对应sheet的索引，新增一个该类型的sheet在上一条记录上的SheetIndex加1作为本条记录的值
        /// </summary>
        public int SheetIndex { get; set; }

        /// <summary>
        /// 当前Sheet内数据行数
        /// </summary>
        public int DataRowCount => this._dataRowCounter;

        /// <summary>
        /// 当前Sheet内数据行数，此处默认为1，
        /// 是因为创建Sheet后必然包含表头会占用一行
        /// </summary>
        private int _dataRowCounter = 1;

        /// <summary>
        /// 递增追加的数据行
        /// </summary>
        internal void IncrementDataRowCount()
        {
            Interlocked.Increment(ref _dataRowCounter);
        }

        /// <summary>
        /// 构建内部Sheet名称
        /// </summary>
        /// <returns></returns>
        internal static string BuildInnerSheetName(string sheetName, int sheetIndex)
        {
            Inspector.NotNullOrWhiteSpace(sheetName, "Sheet名称为空或无效");
            Inspector.Validation(sheetIndex < 0, "Sheet索引无效");
            return $"{sheetName}_{sheetIndex}{MultiStageExporterConst.INNER_SHEET_CHAR}";
        }

        /// <summary>
        /// 构建内部Sheet名称
        /// </summary>
        /// <returns></returns>
        internal static string BuildRealSheetName(string sheetName, int sheetIndex)
        {
            Inspector.NotNullOrWhiteSpace(sheetName, "Sheet名称为空或无效");
            Inspector.Validation(sheetIndex < 0, "Sheet索引无效");

            return $"{sheetName}_{sheetIndex}";
        }
    }
}
