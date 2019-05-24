using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Spire.Xls;

namespace KeLi.ExcelMerge.App.Assists
{
    /// <summary>
    /// 表格辅助
    /// </summary>
    public static class ExcelAssist
    {
        /// <summary>
        /// 导入到表格控件中
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dgv"></param>
        /// <param name="objs"></param>
        public static void ImportDgv<T>(this DataGridView dgv, List<T> objs)
        {
            if (dgv.ColumnCount == 0)
            {
                for (var i = 0; i < typeof(T).GetProperties().Length; i++)
                {
                    var p = typeof(T).GetProperties()[i];
                    var pDcrp = GetDcrp(p);

                    var column = new DataGridViewTextBoxColumn
                    {
                        Name = p.Name,
                        DataPropertyName = p.Name,
                        HeaderText = string.IsNullOrEmpty(pDcrp) ? null : pDcrp,
                        FillWeight = pDcrp == null || pDcrp.Length > 10 ? 7
                            : pDcrp.Length > 6 ? 4
                            : pDcrp.Length < 4 ? 3 : pDcrp.Length
                    };

                    dgv.Columns.Add(column);
                }
            }

            dgv.DataSource = objs;
            dgv.SetDgvStyle();
        }

        /// <summary>
        /// 导入到表格控件中
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dgv"></param>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        public static List<T> ImportDgv<T>(this DataGridView dgv, string filePath, string sheetName = "Sheet1")
        {
            if (dgv.ColumnCount == 0)
            {
                for (var i = 0; i < typeof(T).GetProperties().Length; i++)
                {
                    var p = typeof(T).GetProperties()[i];
                    var pDcrp = GetDcrp(p);

                    var column = new DataGridViewTextBoxColumn
                    {
                        Name = p.Name,
                        DataPropertyName = p.Name,
                        HeaderText = string.IsNullOrEmpty(pDcrp) ? null : pDcrp,
                        FillWeight = pDcrp == null || pDcrp.Length > 10 ? 7
                            : pDcrp.Length > 6 ? 4
                            : pDcrp.Length < 4 ? 3
                            : pDcrp.Length
                    };

                    dgv.Columns.Add(column);
                }
            }

            var results = ImportData<T>(filePath, sheetName);

            dgv.ImportDgv(results);

            return results;
        }

        /// <summary>
        /// 导入到内存
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static List<T> ImportData<T>(string filePath, string sheetName = null)
        {
            var fileInfo = new FileInfo(filePath);
            var results = new List<T>();

            using (var excel = new ExcelPackage(fileInfo))
            {
                var sheets = excel.Workbook.Worksheets;
                var worksheet = sheetName == null ? sheets.FirstOrDefault() : sheets[sheetName];
                var cells = worksheet?.Cells.Value as object[,];

                if (cells == null)
                    return new List<T>();

                for (var i = 1; i < worksheet.Dimension.Rows; i++)
                {
                    var obj = Activator.CreateInstance<T>();
                    var ps = obj.GetType().GetProperties();

                    for (var j = 0; j < worksheet.Dimension.Columns; j++)
                    {
                        var index = j;
                        var pls = ps.Where(w => GetDcrp(w).Equals(cells[0, index]) || w.Name.Equals(cells[0, index]));

                        foreach (var p in pls)
                        {
                            var val = Convert.ChangeType(cells[i, index], p.PropertyType);

                            p.SetValue(obj, cells[i, index] != DBNull.Value ? val : null, null);
                            break;
                        }
                    }

                    results.Add(obj);
                }
            }

            return results;
        }

        /// <summary>
        /// 导入到内存
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        public static DataTable ImportData(string filePath, string sheetName = null)
        {
            using (var workbook = new Workbook())
            {
                workbook.LoadFromFile(filePath);

                var worksheet = sheetName == null ? workbook.Worksheets.FirstOrDefault() : workbook.Worksheets[sheetName];

                return (worksheet as Worksheet)?.ExportDataTable();
            }
        }

        /// <summary>
        /// 导出到文件
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        public static ExcelPackage ExportFile(this DataGridView dgv, string filePath, string sheetName = "Sheet1")
        {
            var fileInfo = new FileInfo(filePath);
            var excel = new ExcelPackage(fileInfo);

            if (excel.Workbook.Worksheets.FirstOrDefault(f => f.Name == sheetName) != null)
                excel.Workbook.Worksheets.Delete(sheetName);

            var worksheet = excel.Workbook.Worksheets.Add(sheetName);
            var index = 0;

            worksheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            foreach (var column in dgv.Columns.Cast<DataGridViewColumn>().Where(w => w.Visible).ToList())
            {
                worksheet.Cells[1, index + 1].Value = column.HeaderText;
                worksheet.Column(index + 1).Width = column.HeaderText.Length > 10 ? 15
                    : column.HeaderText.Length > 6 ? 20
                    : column.HeaderText.Length < 4 ? 8 : 10;
                index++;
            }

            for (var i = 0; i < dgv.RowCount; i++)
            {
                index = 0;

                foreach (var column in dgv.Columns.Cast<DataGridViewColumn>().Where(w => w.Visible).ToList())
                {
                    // 表格值为空，数据仍然存在需要值的情形
                    var val = dgv.Rows[i].Cells[column.Name].Value;
                    var tag = dgv.Rows[i].Cells[column.Name].Tag;
                    var isNull = string.IsNullOrWhiteSpace(val?.ToString());

                    worksheet.Cells[i + 2, index + 1].Value = isNull ? tag : val;
                    index++;
                }
            }

            worksheet.SetFit();
            excel.Save();

            return excel;
        }

        /// <summary>
        /// 导出到文件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="objs"></param>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        public static ExcelPackage ExportFile<T>(this List<T> objs, string filePath, string sheetName = "Sheet1")
        {
            var fileInfo = new FileInfo(filePath);
            var excel = new ExcelPackage(fileInfo);

            if (excel.Workbook.Worksheets.FirstOrDefault(f => f.Name == sheetName) != null)
                excel.Workbook.Worksheets.Delete(sheetName);

            var worksheet = excel.Workbook.Worksheets.Add(sheetName);
            var index = 0;

            foreach (var p in typeof(T).GetProperties())
            {
                worksheet.Cells[1, index + 1].Value = GetDcrp(p);
                index++;
            }

            for (var i = 0; i < objs.Count; i++)
            {
                index = 0;

                foreach (var p in typeof(T).GetProperties())
                {
                    worksheet.Cells[i + 2, index + 1].Value = p.GetValue(objs[i]);
                    index++;
                }
            }

            worksheet.SetFit();
            excel.Save();

            return excel;
        }

        /// <summary>
        /// 导出到文件
        /// </summary>
        /// <param name="table"></param>
        /// <param name="filePath"></param>
        /// <param name="sheetName"></param>
        public static ExcelPackage ExportFile(this DataTable table, string filePath, string sheetName = "Sheet1")
        {
            var fileInfo = new FileInfo(filePath);
            var excel = new ExcelPackage(fileInfo);

            if (excel.Workbook.Worksheets.FirstOrDefault(f => f.Name == sheetName) != null)
                excel.Workbook.Worksheets.Delete(sheetName);

            var worksheet = excel.Workbook.Worksheets.Add(sheetName);
            var index = 0;

            foreach (var column in table.Columns.Cast<DataColumn>().ToList())
            {
                worksheet.Cells[1, index + 1].Value = column.ColumnName;
                index++;
            }

            for (var i = 0; i < table.Rows.Count; i++)
            {
                index = 0;

                foreach (var column in table.Columns.Cast<DataColumn>().ToList())
                {
                    worksheet.Cells[i + 2, index + 1].Value = table.Rows[i][column.ColumnName];
                    index++;
                }
            }

            worksheet.SetFit();
            excel.Save();

            return excel;
        }

        /// <summary>
        /// 设置自定义样式
        /// </summary>
        /// <param name="excel"></param>
        /// <param name="action"></param>
        /// <param name="sheetName"></param>
        public static void SetExcelStyle(this ExcelPackage excel, Action<ExcelWorksheet> action, string sheetName = "Sheet1")
        {
            action?.Invoke(excel.Workbook.Worksheets[sheetName]);
            excel.Save();
        }

        /// <summary>
        /// 设置表格控件样式
        /// </summary>
        /// <param name="dgv"></param>
        public static void SetDgvStyle(this DataGridView dgv)
        {
            // 背景色
            dgv.BackgroundColor = Color.Gray;
            
            // 无边框样式
            dgv.BorderStyle = BorderStyle.None;
            
            // 禁止用户添加行
            dgv.AllowUserToAddRows = false;

            // 禁止调整行高
            dgv.RowTemplate.Height = 25;

            // 行标题不显示
            dgv.RowHeadersVisible = false;

            // 设置两级标题高度，长标题文字调整标高也是一种较好的解决方式
            dgv.ColumnHeadersHeight = 50;

            // 关闭自动设置标题高度
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            // 标题居中
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            // 内容单元格居中对齐
            dgv.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            if(dgv.ColumnCount < 7)
                // 填充模式
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // 整行选中
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        }

        /// <summary>
        /// 获取属性的注释
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public static string GetDcrp(this PropertyInfo p)
        {
            var objs = p.GetCustomAttributes(typeof(DescriptionAttribute), false);

            // 为了不抛空指针异常，必须返回空字符串
            return objs.Length == 0 ? string.Empty : (objs[0] as DescriptionAttribute)?.Description;
        }

        /// <summary>
        /// 计算值
        /// </summary>
        /// <param name="range"></param>
        /// <param name="formula"></param>
        public static void CalcValue(this ExcelRange range, string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                throw new Exception();

            range.Formula = formula;
        }

        /// <summary>
        /// 设置单元格数值格式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="format"></param>
        public static void SetNumberformat(this ExcelRange range, string format)
        {
            if (string.IsNullOrWhiteSpace(format))
                throw new Exception();

            range.Style.Numberformat.Format = format;
        }

        /// <summary>
        /// 设置单元格水平对齐方式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="hAlign"></param>
        public static void SetAlign(this ExcelRange range, ExcelHorizontalAlignment hAlign)
        {
            range.Style.HorizontalAlignment = hAlign;
        }

        /// <summary>
        /// 设置单元格垂直对齐方式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="vAlign"></param>
        public static void SetAlign(this ExcelRange range, ExcelVerticalAlignment vAlign)
        {
            range.Style.VerticalAlignment = vAlign;
        }

        /// <summary>
        /// 设置单元格对齐方式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="hAlign"></param>
        /// <param name="vAlign"></param>
        public static void SetAlign(this ExcelRange range, ExcelHorizontalAlignment hAlign, ExcelVerticalAlignment vAlign)
        {
            range.Style.HorizontalAlignment = hAlign;
            range.Style.VerticalAlignment = vAlign;
        }

        /// <summary>
        /// 自动换行
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="wrap"></param>
        public static void SetWrapText(this ExcelWorksheet worksheet, bool wrap = true)
        {
            worksheet.Cells.Style.WrapText = wrap;
        }

        /// <summary>
        /// 自动换行
        /// </summary>
        /// <param name="range"></param>
        /// <param name="wrap"></param>
        public static void SetWrapText(this ExcelRange range, bool wrap = true)
        {
            range.Style.WrapText = wrap;
        }

        /// <summary>
        /// 设置单元格字体样式
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="color"></param>
        /// <param name="name"></param>
        /// <param name="size"></param>
        public static void SetFont(this ExcelWorksheet worksheet, Color color, string name, int size = 12)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new Exception();

            worksheet.Cells.SetFont(color, name, size);
        }

        /// <summary>
        /// 设置单元格字体样式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="color"></param>
        /// <param name="name"></param>
        /// <param name="size"></param>
        public static void SetFont(this ExcelRange range, Color color, string name, int size = 12)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new Exception();

            // 字体是否粗体
            range.Style.Font.Bold = true;

            // 字体颜色
            range.Style.Font.Color.SetColor(color);

            // 字体
            range.Style.Font.Name = name;

            // 字体大小
            range.Style.Font.Size = size;
        }

        /// <summary>
        /// 设置单元格填充样式
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="style"></param>
        public static void SetFill(this ExcelWorksheet worksheet, ExcelFillStyle style)
        {
            worksheet.Cells.Style.Fill.PatternType = style;
        }

        /// <summary>
        /// 设置单元格填充样式
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="color"></param>
        public static void SetFill(this ExcelWorksheet worksheet, Color color)
        {
            worksheet.Cells.Style.Fill.BackgroundColor.SetColor(color);
        }

        /// <summary>
        /// 设置单元格填充样式
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        public static void SetFill(this ExcelWorksheet worksheet, ExcelFillStyle style, Color color)
        {
            worksheet.Cells.Style.Fill.PatternType = style;
            worksheet.Cells.Style.Fill.BackgroundColor.SetColor(color);
        }

        /// <summary>
        /// 设置单元格填充样式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="style"></param>
        public static void SetFill(this ExcelRange range, ExcelFillStyle style)
        {
            range.Style.Fill.PatternType = style;
        }

        /// <summary>
        /// 设置单元格填充背景色
        /// </summary>
        /// <param name="range"></param>
        /// <param name="color"></param>
        public static void SetFill(this ExcelRange range, Color color)
        {
            range.Style.Fill.BackgroundColor.SetColor(color);
        }

        /// <summary>
        /// 设置单元格填充样式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        public static void SetFill(this ExcelRange range, ExcelFillStyle style, Color color)
        {
            range.Style.Fill.PatternType = style;
            range.Style.Fill.BackgroundColor.SetColor(color);
        }

        /// <summary>
        /// 设置单元格边框样式
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="style"></param>
        public static void SetBorder(this ExcelWorksheet worksheet, ExcelBorderStyle style)
        {
            worksheet.Cells.Style.Border.BorderAround(style);
        }

        /// <summary>
        /// 设置单元格边框样式
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        public static void SetBorder(this ExcelWorksheet worksheet, ExcelBorderStyle style, Color color)
        {
            worksheet.Cells.Style.Border.BorderAround(style, color);
        }

        /// <summary>
        /// 设置单元格边框样式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="style"></param>
        public static void SetBorder(this ExcelRange range, ExcelBorderStyle style)
        {
            range.Style.Border.BorderAround(style);
        }

        /// <summary>
        /// 设置单元格边框样式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        public static void SetBorder(this ExcelRange range, ExcelBorderStyle style, Color color)
        {
            range.Style.Border.BorderAround(style, color);
        }

        /// <summary>
        /// 单独设置单元格某边边框样式
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="position"></param>
        /// <param name="style"></param>
        public static void SetBorderItem(this ExcelWorksheet worksheet, Position position, ExcelBorderStyle style)
        {
            worksheet.Cells.SetBorderItem(position, style);
        }

        /// <summary>
        /// 单独设置单元格某边边框颜色
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="position"></param>
        /// <param name="color"></param>
        public static void SetBorderItem(this ExcelWorksheet worksheet, Position position, Color color)
        {
            worksheet.Cells.SetBorderItem(position, color);
        }

        /// <summary>
        /// 设置单元格某边边框样式
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="position"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        public static void SetBorderItem(this ExcelWorksheet worksheet, Position position, ExcelBorderStyle style, Color color)
        {
            worksheet.Cells.SetBorderItem(position, style, color);
        }

        /// <summary>
        /// 单独设置单元格某边边框样式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="position"></param>
        /// <param name="style"></param>
        public static void SetBorderItem(this ExcelRange range, Position position, ExcelBorderStyle style)
        {
            switch (position)
            {
                case Position.Left:
                    range.Style.Border.Left.Style = style;
                    break;

                case Position.Right:
                    range.Style.Border.Right.Style = style;
                    break;

                case Position.Top:
                    range.Style.Border.Top.Style = style;
                    break;

                case Position.Bottom:
                    range.Style.Border.Bottom.Style = style;
                    break;

                case Position.Diagonal:
                    range.Style.Border.Diagonal.Style = style;
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(position), position, null);
            }
        }

        /// <summary>
        /// 单独设置单元格某边边框颜色
        /// </summary>
        /// <param name="range"></param>
        /// <param name="position"></param>
        /// <param name="color"></param>
        public static void SetBorderItem(this ExcelRange range, Position position, Color color)
        {
            switch (position)
            {
                case Position.Left:
                    range.Style.Border.Bottom.Color.SetColor(color);
                    break;

                case Position.Right:
                    range.Style.Border.Right.Color.SetColor(color);
                    break;

                case Position.Top:
                    range.Style.Border.Top.Color.SetColor(color);
                    break;

                case Position.Bottom:
                    range.Style.Border.Bottom.Color.SetColor(color);
                    break;

                case Position.Diagonal:
                    range.Style.Border.Bottom.Color.SetColor(color);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(position), position, null);
            }
        }

        /// <summary>
        /// 设置单元格某边边框样式
        /// </summary>
        /// <param name="range"></param>
        /// <param name="position"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        public static void SetBorderItem(this ExcelRange range, Position position, ExcelBorderStyle style, Color color)
        {
            switch (position)
            {
                case Position.Left:
                    range.Style.Border.Left.Style = style;
                    range.Style.Border.Left.Color.SetColor(color);
                    break;

                case Position.Right:
                    range.Style.Border.Right.Style = style;
                    range.Style.Border.Right.Color.SetColor(color);
                    break;

                case Position.Top:
                    range.Style.Border.Top.Style = style;
                    range.Style.Border.Top.Color.SetColor(color);
                    break;

                case Position.Bottom:
                    range.Style.Border.Bottom.Style = style;
                    range.Style.Border.Bottom.Color.SetColor(color);
                    break;

                case Position.Diagonal:
                    range.Style.Border.Diagonal.Style = style;
                    range.Style.Border.Diagonal.Color.SetColor(color);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(position), position, null);
            }
        }

        /// <summary>
        /// 设置单元格文字伸缩适应单元格大小
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="shrink"></param>
        public static void SetShrink(this ExcelWorksheet worksheet, bool shrink = true)
        {
            worksheet.Cells.Style.ShrinkToFit = shrink;
        }

        /// <summary>
        /// 设置单元格文字伸缩适应单元格大小
        /// </summary>
        /// <param name="range"></param>
        /// <param name="shrink"></param>
        public static void SetShrink(this ExcelRange range, bool shrink = true)
        {
            range.Style.ShrinkToFit = shrink;
        }

        /// <summary>
        /// 设置单元格伸缩适应单元格文字长度
        /// </summary>
        /// <param name="worksheet"></param>
        public static void SetFit(this ExcelWorksheet worksheet)
        {
            worksheet.Cells.AutoFitColumns();
        }

        /// <summary>
        /// 设置单元格伸缩适应单元格文字长度
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="width"></param>
        public static void SetFit(this ExcelWorksheet worksheet, double width)
        {
            worksheet.Cells.AutoFitColumns(width);
        }

        /// <summary>
        /// 设置单元格伸缩适应单元格文字长度
        /// </summary>
        /// <param name="range"></param>
        /// <param name="fit"></param>
        public static void SetFit(this ExcelRange range, bool fit = true)
        {
            range.AutoFitColumns();
        }

        /// <summary>
        /// 设置表格列伸缩适应单元格文字长度
        /// </summary>
        /// <param name="column"></param>
        public static void SetFit(this ExcelColumn column)
        {
            column.AutoFit();
        }

        /// <summary>
        /// 设置表格列伸缩适应单元格文字长度
        /// </summary>
        /// <param name="column"></param>
        /// <param name="width"></param>
        public static void SetFit(this ExcelColumn column, double width)
        {
            column.AutoFit(width);
        }

        /// <summary>
        /// 设置默认行高
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="height"></param>
        public static void SetDefaultHeight(this ExcelWorksheet worksheet, int height = 15)
        {
            worksheet.DefaultRowHeight = height;
        }

        /// <summary>
        /// 设置默认列宽
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="width"></param>
        public static void SetDefaultWidth(this ExcelWorksheet worksheet, int width = 15)
        {
            worksheet.DefaultColWidth = width;
        }

        /// <summary>
        /// 设置行高
        /// </summary>
        /// <param name="row"></param>
        /// <param name="height"></param>
        public static void SetRowHeight(this ExcelRow row, int height = 15)
        {
            row.Height = height;
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="col"></param>
        /// <param name="width"></param>
        public static void SetColumnWidth(this ExcelColumn col, int width = 15)
        {
            col.Width = width;
        }

        /// <summary>
        /// 设置背景
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="style"></param>
        /// <param name="img"></param>
        /// <param name="showLines"></param>
        public static void SetBackground(this ExcelWorksheet worksheet, ExcelFillStyle style, Image img = null, bool showLines = false)
        {
            // 设置线条样式
            worksheet.Cells.Style.Fill.PatternType = style;

            // 设置背景图片
            worksheet.BackgroundImage.Image = img;

            // 设置网格线
            worksheet.View.ShowGridLines = showLines;
        }

        /// <summary>
        /// 设置背景
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="color"></param>
        /// <param name="img"></param>
        /// <param name="showLines"></param>
        public static void SetBackground(this ExcelWorksheet worksheet, Color color, Image img = null, bool showLines = false)
        {
            // 设置背景色
            worksheet.Cells.Style.Fill.BackgroundColor.SetColor(color);

            // 设置背景图片
            worksheet.BackgroundImage.Image = img;

            // 设置网格线
            worksheet.View.ShowGridLines = showLines;
        }

        /// <summary>
        /// 设置背景
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="style"></param>
        /// <param name="color"></param>
        /// <param name="img"></param>
        /// <param name="showLines"></param>
        public static void SetBackground(this ExcelWorksheet worksheet, ExcelFillStyle style, Color color,
            Image img = null, bool showLines = false)
        {
            // 设置线条样式
            worksheet.Cells.Style.Fill.PatternType = style;

            // 设置背景色
            worksheet.Cells.Style.Fill.BackgroundColor.SetColor(color);

            // 设置背景图片
            worksheet.BackgroundImage.Image = img;

            // 设置网格线
            worksheet.View.ShowGridLines = showLines;
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="name"></param>
        /// <param name="img"></param>
        /// <param name="px"></param>
        /// <param name="py"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        public static void SetPicture(this ExcelWorksheet worksheet, string name, Image img, int px, int py, int width = 100, int height = 100)
        {
            // 插入图片
            var picture = worksheet.Drawings.AddPicture(name, img);

            // 设置图片的位置
            picture.SetPosition(px, py);

            // 设置图片的大小
            picture.SetSize(width, height);
        }

        /// <summary>
        /// 给图片加超链接
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="link"></param>
        /// <param name="img"></param>
        /// <param name="name"></param>
        public static void SetLink(this ExcelWorksheet worksheet, string link, Image img, string name)
        {
            if (string.IsNullOrWhiteSpace(link))
                throw new Exception();

            if (img == null)
                throw new Exception();

            var linker = new ExcelHyperLink(link, UriKind.Relative);

            worksheet.Drawings.AddPicture(name, img, linker);
        }

        /// <summary>
        /// 给单元格加超链接
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="link"></param>
        public static void SetCellLink(this ExcelWorksheet worksheet, string link)
        {
            if (string.IsNullOrWhiteSpace(link))
                throw new Exception();

            worksheet.Cells.Hyperlink = new ExcelHyperLink(link, UriKind.Relative);
        }

        /// <summary>
        /// 给单元格加超链接
        /// </summary>
        /// <param name="range"></param>
        /// <param name="link"></param>
        public static void SetCellLink(this ExcelRange range, string link)
        {
            if (string.IsNullOrWhiteSpace(link))
                throw new Exception();

            range.Hyperlink = new ExcelHyperLink(link, UriKind.Relative);
        }

        /// <summary>
        /// 隐藏工作表
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="hidden"></param>
        public static void SetHiddent(this ExcelWorksheet worksheet, bool hidden = true)
        {
            worksheet.Hidden = hidden ? eWorkSheetHidden.Hidden : eWorkSheetHidden.Visible;
        }

        /// <summary>
        /// 隐藏某一行
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="row"></param>
        /// <param name="hidden"></param>
        public static void SetRowHiddent(this ExcelWorksheet worksheet, int row, bool hidden = true)
        {
            worksheet.Row(1).Hidden = hidden;
        }

        /// <summary>
        /// 隐藏某一列
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="col"></param>
        /// <param name="hidden"></param>
        public static void SetColHiddent(this ExcelWorksheet worksheet, int col, bool hidden = true)
        {
            worksheet.Column(col).Hidden = hidden;
        }
    }

    /// <summary>
    /// 位置枚举
    /// </summary>
    public enum Position
    {
        /// <summary>
        /// 左
        /// </summary>
        Left = 0,

        /// <summary>
        /// 右
        /// </summary>
        Right = 1,

        /// <summary>
        /// 顶
        /// </summary>
        Top = 2,

        /// <summary>
        /// 底
        /// </summary>
        Bottom = 3,

        /// <summary>
        /// 对角
        /// </summary>
        Diagonal = 4
    }
}