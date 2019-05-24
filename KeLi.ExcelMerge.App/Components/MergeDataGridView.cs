using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace KeLi.ExcelMerge.App.Components
{
    /// <summary>
    /// 可合并单元格列表控件
    /// </summary>
    public partial class MergeDataGridView : DataGridView
    {
        /// <summary>
        /// 待合并列字段列表
        /// </summary>
        [Browsable(false)]
        public List<string> MergeColumnNames { get; set; } = new List<string>();

        /// <summary>
        /// 列标题合并字典
        /// </summary>
        private readonly Dictionary<int, SpanInfo> _spanRows = new Dictionary<int, SpanInfo>();

        /// <summary>
        /// 单元格信息列表
        /// </summary>
        private readonly List<CellInfo> _cellInfos = new List<CellInfo>();

        /// <summary>
        /// 是否已加载控件初始化设置
        /// </summary>
        private bool _loaded;

        /// <summary>
        /// 默认加载时列的总宽度
        /// </summary>
        private int _sumWidth;

        /// <summary>
        /// 初始化
        /// </summary>
        public MergeDataGridView()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 滚动事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnScroll(object sender, ScrollEventArgs e)
        {
            Refresh();
        }

        /// <summary>
        /// 大小更改事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnSizeChange(object sender, EventArgs e)
        {
            var sumWidth = Columns.Cast<DataGridViewColumn>()
                .Where(w => w.Visible)
                .Select(s => s.Width).Sum();

            if (_sumWidth == 0)
                _sumWidth = sumWidth;

            // 控件变的足够宽，自动填充，防止调整大小时内部机制导致始终两者相差几个像素
            if ( Width - sumWidth > 5 && _loaded)
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // 控件变的足够窄，恢复滚动条
            else if (sumWidth < _sumWidth && _loaded)
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            Refresh();
        }

        /// <summary>
        /// 合并列
        /// </summary>
        /// <param name="headerText">合并后的单元格的列标题</param>
        /// <param name="colIndex">列的索引</param>
        /// <param name="colCount">需要合并的列数</param>
        public void AddSpanHeader(string headerText, int colIndex, int colCount)
        {
            var rightIndex = colIndex + colCount - 1;

            _spanRows[colIndex] = new SpanInfo(headerText, colIndex, rightIndex);
            _spanRows[rightIndex] = new SpanInfo(headerText, colIndex, rightIndex);

            for (var i = colIndex + 1; i < rightIndex; i++)
                _spanRows[i] = new SpanInfo(headerText, colIndex, rightIndex);
        }

        /// <summary>
        /// 设置单元格信息列表
        /// </summary>
        public void SetCellInfos()
        {
            for (var i = 0; i < Columns.Count; i++)
            {
                for (var j = 0; j < Rows.Count; j++)
                {
                    var cellInfo = new CellInfo
                    {
                        RowIndex = j,
                        ColumnIndex = i
                    };

                    var cellVal = Rows[j].Cells[i].Value?.ToString();

                    // 朝上比较
                    for (var k = j; k >= 0; k--)
                    {
                        var tempVal = Rows[k].Cells[i].Value?.ToString();

                        if (tempVal != cellVal)
                            break;

                        cellInfo.UpRowNum++;
                    }

                    // 朝下比较
                    for (var k = j; k < Rows.Count; k++)
                    {
                        var tempVal = Rows[k].Cells[i].Value?.ToString();

                        if (tempVal != cellVal || tempVal == null)
                            break;

                        cellInfo.DownRowNum++;
                    }

                    _cellInfos.Add(cellInfo);
                }
            }
        }

        /// <summary>
        /// 获取左侧最小上方值相等单元格数
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public int GetUpRowNum(int rowIndex, int columnIndex)
        {
            return _cellInfos.Where(w => w.RowIndex == rowIndex && w.ColumnIndex <= columnIndex)
                .Select(s => s.UpRowNum).Min();
        }

        /// <summary>
        /// 获取左侧最小下方值相等单元格数
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public int GetDownRowNum(int rowIndex, int columnIndex)
        {
            return _cellInfos.Where(w => w.RowIndex == rowIndex && w.ColumnIndex <= columnIndex)
                .Select(s => s.DownRowNum).Min();
        }

        /// <summary>
        /// 重绘单元格
        /// </summary>
        /// <param name="e"></param>
        protected override void OnCellPainting(DataGridViewCellPaintingEventArgs e)
        {
            if (!_loaded)
            {
                // 清除选中行
                ClearSelection();

                _loaded = true;
            }

            // 行标题不重写
            if (e.ColumnIndex < 0 || _spanRows.Count == 0)
            {
                base.OnCellPainting(e);
                return;
            }

            // 标题
            if (e.RowIndex == -1)
                DrawTitle(e);

            // 内容
            else if (e.RowIndex > -1 && e.ColumnIndex > -1)
                DrawCell(e);
        }

        /// <summary>
        /// 绘制标题单元格
        /// </summary>
        /// <param name="e"></param>
        private void DrawTitle(DataGridViewCellPaintingEventArgs e)
        {
            e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.Border);

            var g = e.Graphics;
            var rect = e.CellBounds;
            var left = rect.Left;
            var right = rect.Right;
            var top = rect.Top;
            var bottom = rect.Bottom;

            // 网格画笔
            var gridPen = new Pen(GridColor);

            // 背景色画笔
            var backPen = new Pen(DefaultCellStyle.BackColor);

            // 当前一级标题
            var span = _spanRows[e.ColumnIndex];

            // 画中线
            g.DrawLine(span.HeaderText == e.Value?.ToString() ? backPen : gridPen, left - 1, (top + bottom) / 2, right - 1, (top + bottom) / 2);

            // 标题文字
            DrawString(e, span, ref left, ref right, ref top, ref bottom);

            // 画左边线
            g.DrawLine(gridPen, left - 1, top, left - 1, bottom);

            // 画右边线
            g.DrawLine(gridPen, right - 1, top, right - 1, bottom);

            e.Handled = true;
        }

        /// <summary>
        /// 画标题文字
        /// </summary>
        /// <param name="e"></param>
        /// <param name="span"></param>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <param name="top"></param>
        /// <param name="bottom"></param>
        private void DrawString(DataGridViewCellPaintingEventArgs e, SpanInfo span, ref int left, ref int right, ref int top, ref int bottom)
        {
            var g = e.Graphics;
            var rect = e.CellBounds;

            // 网格画笔
            var gridPen = new Pen(GridColor);

            // 填充笔刷
            var fillBrush = new SolidBrush(DefaultCellStyle.BackColor);

            // 标题文字格式
            var sf = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            // 二级标题矩形
            rect = new Rectangle(left, (top + bottom) / 2 + 1, rect.Width, rect.Height / 2);

            // 标题笔刷
            var headerBrush = new SolidBrush(ColumnHeadersDefaultCellStyle.ForeColor);

            g.FillRectangle(fillBrush, left, rect.Top, rect.Width, rect.Height - 2);

            // 一级标题和二级标题不同时，需要画二级标题
            if (span.HeaderText != e.Value?.ToString())
                g.DrawString(e.Value?.ToString(), e.CellStyle.Font, headerBrush, rect, sf);

            // 画分割线
            g.DrawLine(gridPen, right - 1, top, right - 1, bottom);

            left = GetColumnDisplayRectangle(span.LeftIndex, true).Left;

            if (left < 0)
                left = GetCellDisplayRectangle(-1, -1, true).Width;

            right = GetColumnDisplayRectangle(span.RightIndex, true).Right;

            if (right < 0)
                right = rect.Width;

            // 一级标题矩形
            rect = new Rectangle(left, top, right - left, (bottom - top) / 2);

            // 画上半部分底色
            g.FillRectangle(fillBrush, left, top, rect.Width, rect.Height);

            if (span.HeaderText == e.Value?.ToString())
                rect = new Rectangle(left, top, right - left, bottom - top);

            // 始终都需要画一级标题
            g.DrawString(span.HeaderText, e.CellStyle.Font, headerBrush, rect, sf);
        }

        /// <summary>
        /// 绘制内容单元格
        /// </summary>
        /// <param name="e"></param>
        private void DrawCell(DataGridViewCellPaintingEventArgs e)
        {
            // 画到系统自增的行，跳过该重写方法
            if (e.Value == null)
                return;

            if (e.CellStyle.Alignment == DataGridViewContentAlignment.NotSet)
                e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            if (!MergeColumnNames.Contains(Columns[e.ColumnIndex].Name))
                return;

            var rect = e.CellBounds;
            var g = e.Graphics;

            var upRowNum = GetUpRowNum(e.RowIndex, e.ColumnIndex);
            var downRowNum = GetDownRowNum(e.RowIndex, e.ColumnIndex);
            var tag = Columns[e.ColumnIndex].Tag.ToString();

            if (tag != string.Empty)
            {
                var index = Columns[tag]?.Index;

                upRowNum = GetUpRowNum(e.RowIndex, index ?? 0);
                downRowNum = GetDownRowNum(e.RowIndex, index ?? 0);
            }

            var totalRowNum = upRowNum + downRowNum - 1;

            if (totalRowNum < 2)
                return;

            var backBrush = new SolidBrush(e.CellStyle.BackColor);

            if (Rows[e.RowIndex].Selected)
                backBrush.Color = e.CellStyle.SelectionBackColor;

            // 以背景色填充
            g.FillRectangle(backBrush, rect);

            // 画字符串
            PaintingFont(e, rect.Width, upRowNum, downRowNum, totalRowNum);

            // 网格画笔
            var gridPen = new Pen(GridColor);

            // 画下边线
            if (downRowNum == 1)
                g.DrawLine(gridPen, rect.Left, rect.Bottom - 1, rect.Right, rect.Bottom - 1);

            // 画左边线，内容格画线效果有问题，从效果上拟合
            if (rect.Left < 10)
                g.DrawLine(gridPen, rect.Left, rect.Top, rect.Left, rect.Bottom);

            // 画右边线
            g.DrawLine(gridPen, rect.Right - 1, rect.Top, rect.Right - 1, rect.Bottom);

            e.Handled = true;
        }

        /// <summary>
        /// 绘制文字
        /// </summary>
        /// <param name="e"></param>
        /// <param name="cellwidth"></param>
        /// <param name="upRows"></param>
        /// <param name="downRowNum"></param>
        /// <param name="count"></param>
        private static void PaintingFont(DataGridViewCellPaintingEventArgs e, int cellwidth, int upRows, int downRowNum, int count)
        {
            var font = e.CellStyle.Font;
            var g = e.Graphics;
            var fontBrush = new SolidBrush(e.CellStyle.ForeColor);
            var fontheight = (int)g.MeasureString(e.Value?.ToString(), font).Height;
            var fontwidth = (int)g.MeasureString(e.Value?.ToString(), font).Width;
            var cellRect = e.CellBounds;
            var rectX = cellRect.X;
            var rectY = cellRect.Y;
            var rectHeight = cellRect.Height;
            var width = cellwidth - fontwidth;
            var val = e.Value?.ToString();

            switch (e.CellStyle.Alignment)
            {
                case DataGridViewContentAlignment.BottomCenter:
                    g.DrawString(val, font, fontBrush,
                        rectX + width / 2, rectY + rectHeight * downRowNum - fontheight);
                    break;
                case DataGridViewContentAlignment.BottomLeft:
                    g.DrawString(val, font, fontBrush,
                        rectX, rectY + rectHeight * downRowNum - fontheight);
                    break;
                case DataGridViewContentAlignment.BottomRight:
                    g.DrawString(val, font, fontBrush,
                        rectX + width, rectY + rectHeight * downRowNum - fontheight);
                    break;
                case DataGridViewContentAlignment.MiddleCenter:
                    g.DrawString(val, font, fontBrush,
                        rectX + width / 2, rectY - rectHeight * (upRows - 1) + (rectHeight * count - fontheight) / 2);
                    break;
                case DataGridViewContentAlignment.MiddleLeft:
                    g.DrawString(val, font, fontBrush,
                        rectX, rectY - rectHeight * (upRows - 1) + (rectHeight * count - fontheight) / 2);
                    break;
                case DataGridViewContentAlignment.MiddleRight:
                    g.DrawString(val, font, fontBrush,
                        rectX + width, rectY - rectHeight * (upRows - 1) + (rectHeight * count - fontheight) / 2);
                    break;
                case DataGridViewContentAlignment.TopCenter:
                    g.DrawString(val, font, fontBrush,
                        rectX + width / 2, rectY - rectHeight * (upRows - 1));
                    break;
                case DataGridViewContentAlignment.TopLeft:
                    g.DrawString(val, font, fontBrush,
                        rectX, rectY - rectHeight * (upRows - 1));
                    break;
                case DataGridViewContentAlignment.TopRight:
                    g.DrawString(val, font, fontBrush,
                        rectX + width, rectY - rectHeight * (upRows - 1));
                    break;
                default:
                    g.DrawString(val, font, fontBrush,
                        rectX + width / 2, rectY - rectHeight * (upRows - 1) + (rectHeight * count - fontheight) / 2);
                    break;
            }
        }

        /// <summary>
        /// 表头信息
        /// </summary>
        private struct SpanInfo
        {
            /// <summary>
            /// 初始化
            /// </summary>
            /// <param name="headerText"></param>
            /// <param name="leftIndex"></param>
            /// <param name="rightIndex"></param>
            public SpanInfo(string headerText, int leftIndex, int rightIndex)
            {
                HeaderText = headerText;
                LeftIndex = leftIndex;
                RightIndex = rightIndex;
            }

            /// <summary>
            /// 合并后的单元格的列标题
            /// </summary>
            public string HeaderText { get; }

            /// <summary>
            /// 左侧列标题索引值
            /// </summary>
            public int LeftIndex { get; }

            /// <summary>
            /// 右侧列标题索引值
            /// </summary>
            public int RightIndex { get; }
        }

        /// <summary>
        /// 单元格信息
        /// </summary>
        private class CellInfo
        {
            /// <summary>
            /// 行索引
            /// </summary>
            public int RowIndex { get; set; }

            /// <summary>
            /// 列索引
            /// </summary>
            public int ColumnIndex { get; set; }

            /// <summary>
            /// 上面行数
            /// </summary>
            public int UpRowNum { get; set; }

            /// <summary>
            /// 下面行数
            /// </summary>
            public int DownRowNum { get; set; }
        }
    }
}