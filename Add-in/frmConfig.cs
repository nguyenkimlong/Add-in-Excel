using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace Add_in
{
    [ComVisible(true)]
    public class frmConfig : UserControl
    {
        private GroupBox groupBox1;
        private TextBox txtdencot;
        private TextBox txttucot;
        private Label label3;
        private Label label2;
        private Button btnluu;

        public frmConfig()
        {
            InitializeComponent();
        }
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtdencot = new System.Windows.Forms.TextBox();
            this.txttucot = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnluu = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtdencot);
            this.groupBox1.Controls.Add(this.txttucot);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(30, 14);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(193, 100);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Cài đặt";
            // 
            // txtdencot
            // 
            this.txtdencot.Location = new System.Drawing.Point(71, 57);
            this.txtdencot.Name = "txtdencot";
            this.txtdencot.Size = new System.Drawing.Size(98, 20);
            this.txtdencot.TabIndex = 4;
            // 
            // txttucot
            // 
            this.txttucot.Location = new System.Drawing.Point(71, 24);
            this.txttucot.Name = "txttucot";
            this.txttucot.Size = new System.Drawing.Size(98, 20);
            this.txttucot.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 64);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Đến cột :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 31);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(44, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Từ cột :";
            // 
            // btnluu
            // 
            this.btnluu.Location = new System.Drawing.Point(148, 120);
            this.btnluu.Name = "btnluu";
            this.btnluu.Size = new System.Drawing.Size(75, 23);
            this.btnluu.TabIndex = 2;
            this.btnluu.Text = "Lưu";
            this.btnluu.UseVisualStyleBackColor = true;
            this.btnluu.Click += new System.EventHandler(this.btnluu_Click_1);
            // 
            // frmConfig
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnluu);
            this.Controls.Add(this.groupBox1);
            this.Name = "frmConfig";
            this.Size = new System.Drawing.Size(256, 154);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }


        private void btnluu_Click_1(object sender, EventArgs e)
        {
            Setting.FromCol = txttucot.Text.ToUpper();
            Setting.ToCol = txtdencot.Text.ToUpper();
        }
    }
}
