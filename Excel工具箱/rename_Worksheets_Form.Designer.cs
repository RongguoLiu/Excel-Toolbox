namespace Excel工具箱
{
    partial class rename_Worksheets_Form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.IsValueRelated = new System.Windows.Forms.CheckBox();
            this.begin_Rename = new System.Windows.Forms.Button();
            this.NewName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.insert_CellValue = new System.Windows.Forms.Button();
            this.insert_NumA = new System.Windows.Forms.Button();
            this.insert_NumC = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // IsValueRelated
            // 
            this.IsValueRelated.AutoSize = true;
            this.IsValueRelated.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.IsValueRelated.Location = new System.Drawing.Point(104, 69);
            this.IsValueRelated.Name = "IsValueRelated";
            this.IsValueRelated.Size = new System.Drawing.Size(334, 35);
            this.IsValueRelated.TabIndex = 0;
            this.IsValueRelated.Text = "新工作表名与单元格值相关";
            this.IsValueRelated.UseVisualStyleBackColor = true;
            this.IsValueRelated.CheckedChanged += new System.EventHandler(this.IsValueRelated_CheckedChanged);
            // 
            // begin_Rename
            // 
            this.begin_Rename.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.begin_Rename.Location = new System.Drawing.Point(372, 177);
            this.begin_Rename.Name = "begin_Rename";
            this.begin_Rename.Size = new System.Drawing.Size(188, 48);
            this.begin_Rename.TabIndex = 1;
            this.begin_Rename.Text = "开始重命名";
            this.begin_Rename.UseVisualStyleBackColor = true;
            this.begin_Rename.Click += new System.EventHandler(this.begin_Rename_Click);
            // 
            // NewName
            // 
            this.NewName.Location = new System.Drawing.Point(104, 12);
            this.NewName.Name = "NewName";
            this.NewName.Size = new System.Drawing.Size(456, 35);
            this.NewName.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(12, 15);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 31);
            this.label1.TabIndex = 3;
            this.label1.Text = "新表名";
            // 
            // insert_CellValue
            // 
            this.insert_CellValue.Enabled = false;
            this.insert_CellValue.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.insert_CellValue.Location = new System.Drawing.Point(104, 110);
            this.insert_CellValue.Name = "insert_CellValue";
            this.insert_CellValue.Size = new System.Drawing.Size(148, 48);
            this.insert_CellValue.TabIndex = 4;
            this.insert_CellValue.Text = "单元格值";
            this.insert_CellValue.UseVisualStyleBackColor = true;
            this.insert_CellValue.Click += new System.EventHandler(this.insert_CellValue_Click);
            // 
            // insert_NumA
            // 
            this.insert_NumA.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.insert_NumA.Location = new System.Drawing.Point(258, 110);
            this.insert_NumA.Name = "insert_NumA";
            this.insert_NumA.Size = new System.Drawing.Size(148, 48);
            this.insert_NumA.TabIndex = 5;
            this.insert_NumA.Text = "序号(1)";
            this.insert_NumA.UseVisualStyleBackColor = true;
            this.insert_NumA.Click += new System.EventHandler(this.insert_NumA_Click);
            // 
            // insert_NumC
            // 
            this.insert_NumC.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.insert_NumC.Location = new System.Drawing.Point(412, 110);
            this.insert_NumC.Name = "insert_NumC";
            this.insert_NumC.Size = new System.Drawing.Size(148, 48);
            this.insert_NumC.TabIndex = 6;
            this.insert_NumC.Text = "序号(一)";
            this.insert_NumC.UseVisualStyleBackColor = true;
            this.insert_NumC.Click += new System.EventHandler(this.insert_NumC_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(12, 70);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 31);
            this.label2.TabIndex = 7;
            this.label2.Text = "插入";
            // 
            // rename_Worksheets_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(578, 237);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.insert_NumC);
            this.Controls.Add(this.insert_NumA);
            this.Controls.Add(this.insert_CellValue);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.NewName);
            this.Controls.Add(this.begin_Rename);
            this.Controls.Add(this.IsValueRelated);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "rename_Worksheets_Form";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "批量重命名工作表";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox IsValueRelated;
        private System.Windows.Forms.Button begin_Rename;
        private System.Windows.Forms.TextBox NewName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button insert_NumA;
        private System.Windows.Forms.Button insert_NumC;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button insert_CellValue;
    }
}