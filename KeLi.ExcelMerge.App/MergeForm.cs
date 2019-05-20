using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace KeLi.ExcelMerge.App
{
    public partial class MergeForm : Form
    {
        private readonly List<AreaKpi> _spaces = new List<AreaKpi>();

        public MergeForm()
        {
            InitializeComponent();

            _spaces.AddRange(dgvFile1.ImportDgv<AreaKpi>(@"E:\My Unfiled\Test1.xlsx"));
            _spaces.AddRange(dgvFile2.ImportDgv<AreaKpi>(@"E:\My Unfiled\Test2.xlsx"));
        }

        private void MergeForm_Load(object sender, EventArgs e)
        {
            dgvFile1.ClearSelection();
            dgvFile2.ClearSelection();

            _spaces.ExportFile(@"E:\My Unfiled\Test3.xlsx");
        }
    }
}
