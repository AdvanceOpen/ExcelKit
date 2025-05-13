using System;
using System.Text;
using System.IO;
using System.Collections.Generic;

namespace ExcelKit.Core.ExcelRead
{
    /// <summary>
    /// Excel读取
    /// </summary>
    public interface IReadExcelContext
    {
        /// <summary>
        /// 获取所有Sheet的名称
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>Sheet名称集</returns>
        IEnumerable<string> GetSheetNames(string filePath);

        /// <summary>
        /// 获取所有Sheet的名称
        /// </summary>
        /// <param name="stream">文件流</param>
        /// <param name="disposeStream">是否释放流，默认释放</param>
        /// <returns>Sheet名称集</returns>
        IEnumerable<string> GetSheetNames(Stream stream, bool disposeStream = true);

        /// <summary>
        /// 读取行数据
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="option">读取配置项</param>
        void ReadRows(string filePath, ReadRowsOptions option);

        /// <summary>
        /// 读取行数据
        /// </summary>
        /// <param name="stream">Excel文件流</param>
        /// <param name="option">读取配置项</param>
        void ReadRows(Stream stream, ReadRowsOptions option);

        /// <summary>
        /// 读取Sheet中的数据
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="option">读取配置项</param>
        void ReadSheet(string filePath, ReadSheetDicOptions option);

        /// <summary>
        /// 读取Sheet中的数据
        /// </summary>
        /// <param name="stream">Excel文件流</param>
        /// <param name="option">读取配置项</param>
        void ReadSheet(Stream stream, ReadSheetDicOptions option);

        /// <summary>
        /// 读取Sheet中的数据
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="option">读取配置项</param>
        void ReadSheet<T>(string filePath, ReadSheetOptions<T> option) where T : class, new();

        /// <summary>
        /// 读取Sheet中的数据
        /// </summary>
        /// <param name="stream">Excel文件流</param>
        /// <param name="option">读取配置项</param>
        void ReadSheet<T>(Stream stream, ReadSheetOptions<T> option) where T : class, new();

        /// <summary>
        /// 读取Sheet行总数
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="option">读取配置项</param>
        int ReadSheetRowsCount(string filePath, ReadSheetRowsCountOptions option);

        /// <summary>
        /// 读取Sheet行总数
        /// </summary>
        /// <param name="stream">Excel文件流</param>
        /// <param name="option">读取配置项</param>
        int ReadSheetRowsCount(Stream stream, ReadSheetRowsCountOptions option);
    }
}
