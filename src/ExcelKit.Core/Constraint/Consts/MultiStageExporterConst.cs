using System;
using System.Text;
using System.Collections.Generic;

namespace ExcelKit.Core.Constraint.Consts
{
    /// <summary>
    /// 导出器常量约束
    /// </summary>
    internal class MultiStageExporterConst
    {
        /// <summary>
        /// 内部创建Sheet时，Sheet名称包含的字符
        /// </summary>
        internal const string INNER_SHEET_CHAR = "@%&";

        /// <summary>
        /// 单Sheet最大的数据行数
        /// </summary>
        internal const uint SheetMaxRowCount = 1048200;
    }
}
