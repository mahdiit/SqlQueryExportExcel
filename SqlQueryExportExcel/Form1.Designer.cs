
namespace SqlQueryExportExcel
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtUsername = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtIP = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.chkWinAthu = new System.Windows.Forms.CheckBox();
            this.txtQuery = new System.Windows.Forms.TextBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.exportDialog = new System.Windows.Forms.SaveFileDialog();
            this.txtDb = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtUsername
            // 
            this.txtUsername.Location = new System.Drawing.Point(12, 27);
            this.txtUsername.Name = "txtUsername";
            this.txtUsername.Size = new System.Drawing.Size(117, 22);
            this.txtUsername.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 14);
            this.label1.TabIndex = 1;
            this.label1.Text = "Username";
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(135, 29);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(113, 22);
            this.txtPassword.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(135, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 14);
            this.label2.TabIndex = 1;
            this.label2.Text = "Password";
            // 
            // txtIP
            // 
            this.txtIP.Location = new System.Drawing.Point(12, 74);
            this.txtIP.Name = "txtIP";
            this.txtIP.Size = new System.Drawing.Size(190, 22);
            this.txtIP.TabIndex = 0;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(14, 57);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(42, 14);
            this.label3.TabIndex = 1;
            this.label3.Text = "Server";
            // 
            // chkWinAthu
            // 
            this.chkWinAthu.AutoSize = true;
            this.chkWinAthu.Location = new System.Drawing.Point(209, 76);
            this.chkWinAthu.Name = "chkWinAthu";
            this.chkWinAthu.Size = new System.Drawing.Size(154, 18);
            this.chkWinAthu.TabIndex = 2;
            this.chkWinAthu.Text = "Windows Athuntication";
            this.chkWinAthu.UseVisualStyleBackColor = true;
            // 
            // txtQuery
            // 
            this.txtQuery.Location = new System.Drawing.Point(12, 115);
            this.txtQuery.Multiline = true;
            this.txtQuery.Name = "txtQuery";
            this.txtQuery.Size = new System.Drawing.Size(382, 220);
            this.txtQuery.TabIndex = 3;
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(294, 341);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(100, 23);
            this.btnExport.TabIndex = 4;
            this.btnExport.Text = "Export Data";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // exportDialog
            // 
            this.exportDialog.Filter = "Excel File|*.xlsx";
            // 
            // txtDb
            // 
            this.txtDb.Location = new System.Drawing.Point(254, 29);
            this.txtDb.Name = "txtDb";
            this.txtDb.Size = new System.Drawing.Size(113, 22);
            this.txtDb.TabIndex = 0;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(254, 12);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(22, 14);
            this.label4.TabIndex = 1;
            this.label4.Text = "DB";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(406, 378);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.txtQuery);
            this.Controls.Add(this.chkWinAthu);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtDb);
            this.Controls.Add(this.txtPassword);
            this.Controls.Add(this.txtIP);
            this.Controls.Add(this.txtUsername);
            this.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Export Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtUsername;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtIP;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox chkWinAthu;
        private System.Windows.Forms.TextBox txtQuery;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.SaveFileDialog exportDialog;
        private System.Windows.Forms.TextBox txtDb;
        private System.Windows.Forms.Label label4;
    }
}

