using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace KeLi.ExcelMerge.App
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
        /// 标题垂直居中
        /// </summary>
        public void LoadDgv()
        {
            var data = new List<TestSecond>
            {
                new TestSecond("1","2","3","4","5","6","7","8","9","22","33","44","55","22","33","44","55","22","33","44","55","22","33","44","55"),
                new TestSecond("1","2","3","4","5","6","7","8","59","22","33","44","55","22","33","44","55","22","33","44","55","22","33","44","55"),
                new TestSecond("1","3","3","4","53","6","7","8","59","22","33","44","55","22","33","44","55","22","33","44","55","22","33","44","55"),
                new TestSecond("2","2","3","4","5","6","77","8","9","22","33","44","55","22","33","44","55","22","33","44","55","22","33","44","55"),
                new TestSecond("2","2","3","4","53","6","7","8","9","22","33","44","55","22","33","44","55","22","33","44","55","22","33","44","55"),
                new TestSecond("3","2","3","4","53","6","77","8","9","22","33","44","55","22","33","44","55","22","33","44","55","22","33","44","55"),
                new TestSecond("3","2","3","4","5","6","77","8","9","22","33","44","55","22","33","44","55","22","33","44","55","22","33","44","55")
            };

            mdgvTest.ImportDgv<TestFirst, TestSecond>(data);
            //mdgvTest.ColumnHeadersDefaultCellStyle.ForeColor= Color.Blue;
            //mdgvTest.DefaultCellStyle.BackColor = Color.BlanchedAlmond;

            mdgvTest.ExportFile<TestFirst, TestSecond>(@"C:\Users\KeLi\Desktop\TestSecond.xlsx");
        }
    }
}
