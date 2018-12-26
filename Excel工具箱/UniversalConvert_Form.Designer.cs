namespace Excel工具箱
{
    partial class UniversalConvert_Form
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
            this.OfficeApplicationTab = new System.Windows.Forms.TabControl();
            this.PowerPointTab = new System.Windows.Forms.TabPage();
            this.PowerPointToSelection = new System.Windows.Forms.Button();
            this.PowerPointToPDF = new System.Windows.Forms.Button();
            this.PowerPointSelection = new System.Windows.Forms.ComboBox();
            this.WordTab = new System.Windows.Forms.TabPage();
            this.OfficeApplicationTab.SuspendLayout();
            this.PowerPointTab.SuspendLayout();
            this.SuspendLayout();
            // 
            // OfficeApplicationTab
            // 
            this.OfficeApplicationTab.Controls.Add(this.PowerPointTab);
            this.OfficeApplicationTab.Controls.Add(this.WordTab);
            this.OfficeApplicationTab.Location = new System.Drawing.Point(12, 12);
            this.OfficeApplicationTab.Name = "OfficeApplicationTab";
            this.OfficeApplicationTab.SelectedIndex = 0;
            this.OfficeApplicationTab.Size = new System.Drawing.Size(488, 184);
            this.OfficeApplicationTab.TabIndex = 0;
            // 
            // PowerPointTab
            // 
            this.PowerPointTab.Controls.Add(this.PowerPointToSelection);
            this.PowerPointTab.Controls.Add(this.PowerPointToPDF);
            this.PowerPointTab.Controls.Add(this.PowerPointSelection);
            this.PowerPointTab.Location = new System.Drawing.Point(8, 39);
            this.PowerPointTab.Name = "PowerPointTab";
            this.PowerPointTab.Padding = new System.Windows.Forms.Padding(3);
            this.PowerPointTab.Size = new System.Drawing.Size(472, 137);
            this.PowerPointTab.TabIndex = 0;
            this.PowerPointTab.Text = "PowerPoint  ";
            this.PowerPointTab.UseVisualStyleBackColor = true;
            // 
            // PowerPointToSelection
            // 
            this.PowerPointToSelection.Location = new System.Drawing.Point(6, 45);
            this.PowerPointToSelection.Name = "PowerPointToSelection";
            this.PowerPointToSelection.Size = new System.Drawing.Size(460, 36);
            this.PowerPointToSelection.TabIndex = 2;
            this.PowerPointToSelection.Text = "转换为所选格式(开发中)";
            this.PowerPointToSelection.UseVisualStyleBackColor = true;
            // 
            // PowerPointToPDF
            // 
            this.PowerPointToPDF.Location = new System.Drawing.Point(6, 87);
            this.PowerPointToPDF.Name = "PowerPointToPDF";
            this.PowerPointToPDF.Size = new System.Drawing.Size(460, 36);
            this.PowerPointToPDF.TabIndex = 1;
            this.PowerPointToPDF.Text = "转换为PDF";
            this.PowerPointToPDF.UseVisualStyleBackColor = true;
            this.PowerPointToPDF.Click += new System.EventHandler(this.PowerPointToPDF_Click);
            // 
            // PowerPointSelection
            // 
            this.PowerPointSelection.FormattingEnabled = true;
            this.PowerPointSelection.Location = new System.Drawing.Point(6, 6);
            this.PowerPointSelection.Name = "PowerPointSelection";
            this.PowerPointSelection.Size = new System.Drawing.Size(460, 33);
            this.PowerPointSelection.TabIndex = 0;
            // 
            // WordTab
            // 
            this.WordTab.Location = new System.Drawing.Point(8, 39);
            this.WordTab.Name = "WordTab";
            this.WordTab.Padding = new System.Windows.Forms.Padding(3);
            this.WordTab.Size = new System.Drawing.Size(472, 137);
            this.WordTab.TabIndex = 1;
            this.WordTab.Text = "Word  ";
            this.WordTab.UseVisualStyleBackColor = true;
            // 
            // UniversalConvert
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(512, 208);
            this.Controls.Add(this.OfficeApplicationTab);
            this.MaximizeBox = false;
            this.Name = "UniversalConvert";
            this.ShowIcon = false;
            this.Text = "通用Office格式转换";
            this.OfficeApplicationTab.ResumeLayout(false);
            this.PowerPointTab.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl OfficeApplicationTab;
        private System.Windows.Forms.TabPage PowerPointTab;
        private System.Windows.Forms.TabPage WordTab;
        private System.Windows.Forms.Button PowerPointToPDF;
        private System.Windows.Forms.ComboBox PowerPointSelection;
        private System.Windows.Forms.Button PowerPointToSelection;
    }
}