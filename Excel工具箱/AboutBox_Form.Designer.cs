namespace Excel工具箱
{
    partial class AboutBox_Form
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
            this.SuspendLayout();
            // 
            // AboutBox_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(578, 0);
            this.Name = "AboutBox_Form";
            this.Text = "正在加载帮助…";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.AboutBox_Form_FormClosing);
            this.Shown += new System.EventHandler(this.AboutBox_Form_Shown);
            this.ResumeLayout(false);

        }

        #endregion
    }
}