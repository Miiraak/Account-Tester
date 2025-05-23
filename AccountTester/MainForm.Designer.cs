﻿namespace AccountTester
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            buttonStart = new Button();
            richTextBoxLogs = new RichTextBox();
            buttonExportForm = new Button();
            buttonCopier = new Button();
            label2 = new Label();
            SuspendLayout();
            // 
            // buttonStart
            // 
            buttonStart.Location = new Point(13, 352);
            buttonStart.Name = "buttonStart";
            buttonStart.Size = new Size(75, 21);
            buttonStart.TabIndex = 0;
            buttonStart.Text = "Start";
            buttonStart.UseVisualStyleBackColor = true;
            buttonStart.Click += ButtonStart_Click;
            // 
            // richTextBoxLogs
            // 
            richTextBoxLogs.HideSelection = false;
            richTextBoxLogs.Location = new Point(12, 30);
            richTextBoxLogs.Name = "richTextBoxLogs";
            richTextBoxLogs.ReadOnly = true;
            richTextBoxLogs.Size = new Size(304, 316);
            richTextBoxLogs.TabIndex = 4;
            richTextBoxLogs.TabStop = false;
            richTextBoxLogs.Text = "";
            // 
            // buttonExportForm
            // 
            buttonExportForm.Location = new Point(241, 352);
            buttonExportForm.Name = "buttonExportForm";
            buttonExportForm.Size = new Size(75, 23);
            buttonExportForm.TabIndex = 5;
            buttonExportForm.Text = "Export";
            buttonExportForm.UseVisualStyleBackColor = true;
            buttonExportForm.Click += ButtonExport_Click;
            // 
            // buttonCopier
            // 
            buttonCopier.Enabled = false;
            buttonCopier.Location = new Point(127, 352);
            buttonCopier.Name = "buttonCopier";
            buttonCopier.Size = new Size(75, 21);
            buttonCopier.TabIndex = 6;
            buttonCopier.Text = "Copy";
            buttonCopier.UseVisualStyleBackColor = true;
            buttonCopier.Click += ButtonCopier_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Consolas", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label2.Location = new Point(12, 9);
            label2.Name = "label2";
            label2.Size = new Size(56, 18);
            label2.TabIndex = 3;
            label2.Text = "Logs :";
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 14F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            ClientSize = new Size(328, 381);
            Controls.Add(buttonCopier);
            Controls.Add(buttonExportForm);
            Controls.Add(richTextBoxLogs);
            Controls.Add(label2);
            Controls.Add(buttonStart);
            Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "MainForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Account Tester v0.7.8";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button buttonStart;
        private RichTextBox richTextBoxLogs;
        private Button buttonExportForm;
        private Button buttonCopier;
        private Label label2;
    }
}
