using System;

namespace KeLi.ExcelMerge.App
{
    /// <summary>
    /// 参照特性
    /// </summary>
    public class ReferenceAttribute : Attribute
    {
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="columnName"></param>
        public ReferenceAttribute(string columnName)
        {
            ColumnName = columnName;
        }

        /// <summary>
        /// 列名
        /// </summary>
        public string ColumnName { get; set; }
    }
}
