namespace KeLi.ExcelMerge.App
{
    partial class MergeForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.dgvFile1 = new System.Windows.Forms.DataGridView();
            this.dgvFile2 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFile1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFile2)).BeginInit();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.dgvFile1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dgvFile2);
            this.splitContainer1.Size = new System.Drawing.Size(767, 511);
            this.splitContainer1.SplitterDistance = 251;
            this.splitContainer1.SplitterWidth = 5;
            this.splitContainer1.TabIndex = 3;
            // 
            // dgvFile1
            // 
            this.dgvFile1.AllowUserToAddRows = false;
            this.dgvFile1.AllowUserToDeleteRows = false;
            this.dgvFile1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvFile1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvFile1.Location = new System.Drawing.Point(0, 0);
            this.dgvFile1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.dgvFile1.Name = "dgvFile1";
            this.dgvFile1.ReadOnly = true;
            this.dgvFile1.RowTemplate.Height = 23;
            this.dgvFile1.Size = new System.Drawing.Size(767, 251);
            this.dgvFile1.TabIndex = 0;
            // 
            // dgvFile2
            // 
            this.dgvFile2.AllowUserToAddRows = false;
            this.dgvFile2.AllowUserToDeleteRows = false;
            this.dgvFile2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvFile2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvFile2.Location = new System.Drawing.Point(0, 0);
            this.dgvFile2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.dgvFile2.Name = "dgvFile2";
            this.dgvFile2.ReadOnly = true;
            this.dgvFile2.RowTemplate.Height = 23;
            this.dgvFile2.Size = new System.Drawing.Size(767, 255);
            this.dgvFile2.TabIndex = 0;
            // 
            // MergeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 511);
            this.Controls.Add(this.splitContainer1);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "MergeForm";
            this.Text = "合并测试";
            this.Load += new System.EventHandler(this.MergeForm_Load);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvFile1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFile2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridView dgvFile1;
        private System.Windows.Forms.DataGridView dgvFile2;
    }
}

