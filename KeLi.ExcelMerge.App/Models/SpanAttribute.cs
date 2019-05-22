using System;

namespace KeLi.ExcelMerge.App.Models
{
    /// <summary>
    /// 标题跨列特性
    /// </summary>
    public class SpanAttribute : Attribute
    {
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="columnSpan"></param>
        public SpanAttribute(int columnSpan)
        {
            ColumnSpan = columnSpan;
        }

        /// <summary>
        /// 跨列数
        /// </summary>
        public int ColumnSpan { get; set; }
    }
}