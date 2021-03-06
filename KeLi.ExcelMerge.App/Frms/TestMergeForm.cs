﻿using System.Collections.Generic;
using System.Windows.Forms;
using KeLi.ExcelMerge.App.Assists;
using KeLi.ExcelMerge.App.Models;

namespace KeLi.ExcelMerge.App.Frms
{
    /// <summary>
    /// 测试可合并单元格控件
    /// </summary>
    public partial class TestMergeForm : Form
    {
        /// <summary>
        /// 初始化
        /// </summary>
        public TestMergeForm()
        {
            InitializeComponent();
            LoadDgv();
        }

        /// <summary>
        /// 加载控件数据
        /// </summary>
        public void LoadDgv()
        {
            var data = new List<TestSecond>
            {
                new TestSecond("商业-集中式", 1500, 1000, 500, 500, 300, 200, 5, 3, "主卧", 300, 0.5, true, "有"),
                new TestSecond("商业-集中式", 1100, 500, 600, 500, 300, 200, 5, 3, "主卧", 300, 0.5, true, "有"),
                new TestSecond("商业-分布式", 1500, 1000, 500, 500, 300, 200, 5, 3, "主卧", 300, 0.5, true, "有"),
                new TestSecond("商业-集中式", 1500, 1000, 500, 500, 300, 200, 5, 3, "主卧", 300, 0.5, true, "有"),
                new TestSecond("商业-分布式", 1300, 700, 600, 500, 300, 200, 5, 3, "主卧", 300, 0.5, true, "有"),
                new TestSecond("商业-分布式", 1500, 1000, 500, 500, 300, 200, 5, 3, "主卧", 300, 0.5, true, "有"),
                new TestSecond("商业-分布式", 1500, 1000, 500, 500, 300, 200, 5, 3, "主卧", 300, 0.5, true, "有")
            };

            mdgvTest.ImportDgv<TestFirst, TestSecond>(data);
            mdgvTest.ExportFile<TestFirst, TestSecond>(@"C:\Users\KeLi\Desktop\TestSecond.xlsx");
        }
    }
}