using System;
using System.Collections.Generic;
using System.Text;
using ExcelKit.Core.Constraint.Enums;

namespace ExcelKit.Core.Attributes
{
    /// <summary>
    /// 导入导出Attribute
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Class)]
    public class ExcelKitAttribute : Attribute
    {
        /// <summary>
        /// 编码
        /// </summary>
        /// <remarks>
        /// Code可默认不指定，此时使用字段名作为Code
        /// </remarks>
        public string Code { get; set; }

        /// <summary>
        /// 描述(读取和导出时必指定)
        /// </summary>
        public string Desc { get; set; }

        /// <summary>
        /// 是否可空
        /// </summary>
        public bool AllowNull { get; set; }

        /// <summary>
        /// 转换器，需要继承自IExportConverter<T>
        /// </summary>
        public Type Converter { get; set; }

        /// <summary>
        /// 转换器辅助参数
        /// </summary>
        public object ConverterParam { get; set; }

        /// <summary>
        /// 字段排序(默认升序)
        /// </summary>
        public float Sort { get; set; }

        /// <summary>
        /// 列的宽度(不指定的话为默认宽度)
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// 对齐方式
        /// </summary>
        public TextAlign Align { get; set; }

        /// <summary>
        /// 字体颜色
        /// </summary>
        public DefineColor FontColor { get; set; }

        /// <summary>
        /// 前景色
        /// </summary>
        public DefineColor ForegroundColor { get; set; } = DefineColor.None;

        /// <summary>
        /// 表头行冻结
        /// </summary>
        public bool HeadRowFrozen { get; set; }

        /// <summary>
        /// 表头行筛选
        /// </summary>
        public bool HeadRowFilter { get; set; }

        /// <summary>
        /// 是否忽略
        /// </summary>
        public bool IsIgnore { get; set; }

        /// <summary>
        /// 是否仅读取时忽略
        /// </summary>
        public bool IsOnlyIgnoreRead { get; set; }

        /// <summary>
        /// 是否仅导出时忽略
        /// </summary>
        public bool IsOnlyIgnoreWrite { get; set; }
    }
}