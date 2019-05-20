﻿using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace KeLi.ExcelMerge.App
{
    /// <summary>
    /// 表格辅助扩展
    /// </summary>
    public static class ExcelAssistEx
    {
        /// <summary>
        /// 导入到表格控件中
        /// </summary>
        /// <typeparam name="Title"></typeparam>
        /// <typeparam name="Model"></typeparam>
        /// <param name="mdgv"></param>
        /// <param name="objs"></param>
        /// <param name="mergeCell"></param>
        public static void ImportDgv<Title, Model>(this MergeDataGridView mdgv, List<Model> objs, bool mergeCell = true)
        {
            for (var i = 0; i < typeof(Model).GetProperties().Length; i++)
            {
                var p = typeof(Model).GetProperties()[i];
                var pDcrp = GetDcrp(p);

                if (pDcrp == null)
                    continue;

                var column = new DataGridViewTextBoxColumn
                {
                    Name = p.Name,
                    DataPropertyName = p.Name,
                    HeaderText = string.IsNullOrEmpty(pDcrp) ? p.Name : pDcrp,
                    //FillWeight = pDcrp.Length > 10 ? 7 : pDcrp.Length > 6 ? 4 : pDcrp.Length < 4 ? 3 : pDcrp.Length,
                    Tag = GetReference(p)
                };

                mdgv.Columns.Add(column);
                mdgv.MergeColumnNames.Add(p.Name);
            }

            // 数据源
            mdgv.DataSource = objs;

            // 设置表格样式
            SetDgvStyle(mdgv);

            // 防止合并后，依旧选中行或单元格显示
           // mdgv.Enabled = false;

            // 设置跨列合并单元格
            MergeHeaders<Title>(mdgv);

            // 设置合并内容单元格
            MergeCell(mdgv);
        }

        /// <summary>
        /// 导出到文件
        /// </summary>
        /// <param name="mdgv"></param>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        public static ExcelPackage ExportFile<Title, Model>(this MergeDataGridView mdgv, string filePath, string sheetName = "Sheet1")
        {
            var fileInfo = new FileInfo(filePath);
            var excel = new ExcelPackage(fileInfo);

            if (excel.Workbook.Worksheets.FirstOrDefault(f => f.Name == sheetName) != null)
                excel.Workbook.Worksheets.Delete(sheetName);

            var worksheet = excel.Workbook.Worksheets.Add(sheetName);
            var index = 0;
            var lastSum = 1;

            worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // 一级标题
            foreach (var p in typeof(Title).GetProperties())
            {
                var spanNum = GetSpan(p);
                var columnDcrp = GetDcrp(p);

                // 合并格子和索引简单的数学关系可知道需要减1，坐标轴图示即可
                worksheet.Cells[1, lastSum, 1, lastSum + spanNum - 1].Value = columnDcrp;

                worksheet.Column(lastSum).Width = columnDcrp.Length > 10 ? 15
                    : columnDcrp.Length > 6 ? 20
                    : columnDcrp.Length < 4 ? 8 : 10;

                // 只有一个格子不要修改融合属性
                if (lastSum != lastSum + spanNum - 1)
                    worksheet.Cells[1, lastSum, 1, lastSum + spanNum - 1].Merge = true;

                lastSum += spanNum;
            }

            // 二级标题
            foreach (var column in mdgv.Columns.Cast<DataGridViewColumn>().Where(w => w.Visible).ToList())
            {
                worksheet.Cells[2, index + 1].Value = column.HeaderText;
                index++;
            }

            // 标题融合
            for (var i = 0; i < typeof(Model).GetProperties().Length; i++)
            {
                // 融合过的，跳过
                if (worksheet.Cells[1, i + 1].Merge)
                    continue;

                // 值不等，跳过
                if (worksheet.Cells[1, i + 1].Value?.ToString() != worksheet.Cells[2, i + 1].Value?.ToString())
                    continue;

                // 融合
                worksheet.Cells[1, i + 1, 2, i + 1].Merge = true;

                // 垂直居中
                worksheet.Cells[1, i + 1, 2, i + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            // 内容
            for (var i = 0; i < mdgv.RowCount; i++)
            {
                index = 0;

                foreach (var column in mdgv.Columns.Cast<DataGridViewColumn>().Where(w => w.Visible).ToList())
                {
                    worksheet.Cells[i + 3, index + 1].Value = mdgv.Rows[i].Cells[column.Name].Value;
                    index++;
                }
            }

            // 内容合并
            // 遍历列
            for (var i = 1; i <= worksheet.Dimension.Columns; i++)
            {
                // 遍历行
                for (var j = 3; j <= worksheet.Dimension.Rows; j++)
                {
                    var upRowsNum = 0;//mdgv.GetUpRowNum(j-3, i);
                    var downRowNum = 0;//mdgv.GetDownRowNum(j-3, i);
                    var curCell = worksheet.Cells[j, i];

                    if (curCell.Merge)
                        continue;

                    // 单元格朝上比较
                    for (var k = j - 1; k > 2; k--)
                    {
                        var tempCell = worksheet.Cells[k, i];

                        if (tempCell.Merge)
                            break;

                        if (tempCell.Value?.ToString() == curCell.Value?.ToString())
                            upRowsNum++;

                        else
                            break;
                    }

                    // 单元格朝下比较
                    for (var k = j + 1; k <= worksheet.Dimension.Rows; k++)
                    {
                        var tempCell = worksheet.Cells[k, i];

                        if (tempCell.Merge)
                            break;

                        if (tempCell.Value?.ToString() == curCell.Value?.ToString())
                            downRowNum++;

                        else
                            break;
                    }

                    if (worksheet.Cells[j - upRowsNum, i].Merge)
                        continue;

                    // 融合
                    worksheet.Cells[j - upRowsNum, i, j + downRowNum, i].Merge = true;
                }
            }

            excel.Save();

            return excel;
        }

        /// <summary>
        /// 合并列标题
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="mdgv"></param>
        public static void MergeHeaders<T>(this MergeDataGridView mdgv)
        {
            var lastSum = 0;

            for (var i = 0; i < typeof(T).GetProperties().Length; i++)
            {
                var p = typeof(T).GetProperties()[i];
                var spanNum = GetSpan(p);

                mdgv.AddSpanHeader(GetDcrp(p), lastSum, spanNum);
                lastSum += spanNum;
            }
        }

        /// <summary>
        /// 合并内容单元格
        /// </summary>
        public static void MergeCell(this MergeDataGridView mdgv)
        {
            mdgv.SetCellInfos();
        }

        /// <summary>
        /// 获取融合后范围内单元格的值
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public static string GetMegerValue(this ExcelWorksheet worksheet, int row, int column)
        {
            var rangeStr = worksheet.MergedCells[row, column];
            var excelRange = worksheet.Cells;
            var cellVal = excelRange[row, column].Value;

            if (rangeStr == null)
                return cellVal?.ToString();

            var startCell = new ExcelAddress(rangeStr).Start;

            return excelRange[startCell.Row, startCell.Column].Value?.ToString() ?? string.Empty;
        }

        /// <summary>
        /// 设置表格控件样式
        /// </summary>
        /// <param name="dgv"></param>
        public static void SetDgvStyle(DataGridView dgv)
        {
            // 行标题不显示
            dgv.RowHeadersVisible = false;

            // 设置两级标题高度
            dgv.ColumnHeadersHeight = 50;

            // 关闭自动设置标题高度
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            // 标题居中
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            // 内容单元格居中对齐
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            // 填充模式
            //dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // 整行选中
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        /// <summary>
        /// 获取属性的描述
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public static string GetDcrp(PropertyInfo p)
        {
            var objs = p.GetCustomAttributes(typeof(DescriptionAttribute), false);

            // 为了不抛空指针异常，必须返回空字符串
            return objs.Length == 0 ? string.Empty : (objs[0] as DescriptionAttribute)?.Description;
        }

        /// <summary>
        /// 获取属性的跨列数
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public static int GetSpan(PropertyInfo p)
        {
            var objs = p.GetCustomAttributes(typeof(SpanAttribute), false);

            if (objs.Length == 0)
                return 1;

            var attr = objs[0] as SpanAttribute;

            if (attr != null)
                return objs.Length == 0 ? 1 : attr.ColumnSpan;

            return 1;
        }

        /// <summary>
        /// 获取属性的参照列
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public static string GetReference(PropertyInfo p)
        {
            var objs = p.GetCustomAttributes(typeof(ReferenceAttribute), false);

            if (objs.Length == 0)
                return string.Empty;

            var attr = objs[0] as ReferenceAttribute;

            if (attr != null)
                return objs.Length == 0 ? string.Empty : attr.ColumnName;

            return string.Empty;
        }
    }
}