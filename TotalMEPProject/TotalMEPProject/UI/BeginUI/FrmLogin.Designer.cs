﻿namespace TotalMEPProject.UI
{
    partial class FrmLogin
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmLogin));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.rbtnAccount = new System.Windows.Forms.RadioButton();
            this.rbtnLiscense = new System.Windows.Forms.RadioButton();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.label1 = new System.Windows.Forms.Label();
            this.lblUserServer = new System.Windows.Forms.Label();
            this.lblPassword = new System.Windows.Forms.Label();
            this.tbUserServer = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.linkForgot = new System.Windows.Forms.LinkLabel();
            this.tbPassword = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnActive = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 110F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel3, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.lblUserServer, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.lblPassword, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.tbUserServer, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel4, 1, 3);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33334F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33334F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(568, 145);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 4;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 110F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 110F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.Controls.Add(this.rbtnAccount, 0, 0);
            this.tableLayoutPanel3.Controls.Add(this.rbtnLiscense, 1, 0);
            this.tableLayoutPanel3.Controls.Add(this.linkLabel2, 2, 0);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(113, 3);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 1;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel3.Size = new System.Drawing.Size(452, 28);
            this.tableLayoutPanel3.TabIndex = 5;
            // 
            // rbtnAccount
            // 
            this.rbtnAccount.AutoSize = true;
            this.rbtnAccount.Checked = true;
            this.rbtnAccount.Location = new System.Drawing.Point(3, 3);
            this.rbtnAccount.Name = "rbtnAccount";
            this.rbtnAccount.Size = new System.Drawing.Size(65, 17);
            this.rbtnAccount.TabIndex = 0;
            this.rbtnAccount.TabStop = true;
            this.rbtnAccount.Text = "Account";
            this.rbtnAccount.UseVisualStyleBackColor = true;
            this.rbtnAccount.CheckedChanged += new System.EventHandler(this.rbtnAccount_CheckedChanged);
            // 
            // rbtnLiscense
            // 
            this.rbtnLiscense.AutoSize = true;
            this.rbtnLiscense.Location = new System.Drawing.Point(103, 3);
            this.rbtnLiscense.Name = "rbtnLiscense";
            this.rbtnLiscense.Size = new System.Drawing.Size(96, 17);
            this.rbtnLiscense.TabIndex = 2;
            this.rbtnLiscense.TabStop = true;
            this.rbtnLiscense.Text = "License Server";
            this.rbtnLiscense.UseVisualStyleBackColor = true;
            this.rbtnLiscense.CheckedChanged += new System.EventHandler(this.rbtnLiscense_CheckedChanged);
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Location = new System.Drawing.Point(215, 5);
            this.linkLabel2.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(65, 13);
            this.linkLabel2.TabIndex = 3;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "Buy License";
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 5);
            this.label1.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Get License From:";
            // 
            // lblUserServer
            // 
            this.lblUserServer.AutoSize = true;
            this.lblUserServer.Location = new System.Drawing.Point(5, 39);
            this.lblUserServer.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this.lblUserServer.Name = "lblUserServer";
            this.lblUserServer.Size = new System.Drawing.Size(90, 13);
            this.lblUserServer.TabIndex = 1;
            this.lblUserServer.Text = "UserName/Email:";
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSize = true;
            this.lblPassword.Location = new System.Drawing.Point(5, 74);
            this.lblPassword.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(56, 13);
            this.lblPassword.TabIndex = 2;
            this.lblPassword.Text = "Password:";
            // 
            // tbUserServer
            // 
            this.tbUserServer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbUserServer.Location = new System.Drawing.Point(115, 39);
            this.tbUserServer.Margin = new System.Windows.Forms.Padding(5, 5, 5, 0);
            this.tbUserServer.Name = "tbUserServer";
            this.tbUserServer.Size = new System.Drawing.Size(448, 20);
            this.tbUserServer.TabIndex = 3;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 70F));
            this.tableLayoutPanel2.Controls.Add(this.linkForgot, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.tbPassword, 0, 0);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(113, 72);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(452, 29);
            this.tableLayoutPanel2.TabIndex = 4;
            // 
            // linkForgot
            // 
            this.linkForgot.AutoSize = true;
            this.linkForgot.Location = new System.Drawing.Point(387, 5);
            this.linkForgot.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this.linkForgot.Name = "linkForgot";
            this.linkForgot.Size = new System.Drawing.Size(43, 13);
            this.linkForgot.TabIndex = 0;
            this.linkForgot.TabStop = true;
            this.linkForgot.Text = "Forgot?";
            // 
            // tbPassword
            // 
            this.tbPassword.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbPassword.Location = new System.Drawing.Point(3, 3);
            this.tbPassword.Name = "tbPassword";
            this.tbPassword.Size = new System.Drawing.Size(376, 20);
            this.tbPassword.TabIndex = 1;
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.ColumnCount = 3;
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 80F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 80F));
            this.tableLayoutPanel4.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.Controls.Add(this.btnCancel, 1, 0);
            this.tableLayoutPanel4.Controls.Add(this.btnActive, 0, 0);
            this.tableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel4.Location = new System.Drawing.Point(113, 107);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.RowCount = 1;
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel4.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel4.Size = new System.Drawing.Size(452, 35);
            this.tableLayoutPanel4.TabIndex = 6;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnCancel.Location = new System.Drawing.Point(85, 5);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(70, 25);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnActive
            // 
            this.btnActive.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnActive.Location = new System.Drawing.Point(5, 5);
            this.btnActive.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this.btnActive.Name = "btnActive";
            this.btnActive.Size = new System.Drawing.Size(70, 25);
            this.btnActive.TabIndex = 0;
            this.btnActive.Text = "Active";
            this.btnActive.UseVisualStyleBackColor = true;
            this.btnActive.Click += new System.EventHandler(this.btnActive_Click);
            // 
            // FrmLogin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(568, 145);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmLogin";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Total MEP Login";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel3.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.tableLayoutPanel4.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.RadioButton rbtnLiscense;
        private System.Windows.Forms.RadioButton rbtnAccount;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblUserServer;
        private System.Windows.Forms.Label lblPassword;
        private System.Windows.Forms.TextBox tbUserServer;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.LinkLabel linkForgot;
        private System.Windows.Forms.TextBox tbPassword;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel4;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnActive;
        private System.Windows.Forms.LinkLabel linkLabel2;
    }
}