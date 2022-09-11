using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using ExcelKit.Core.Attributes;
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
        public int DataRowCount { get; set; }

        /// <summary>
        /// 当前Sheet内数据行数
        /// </summary>
        private int _dataRowCounter = 0;

        /// <summary>
        /// 递增追加的数据行
        /// </summary>
        internal void IncrementDataRowCount()
        {
            DataRowCount = Interlocked.Increment(ref _dataRowCounter);
        }
    }
}
