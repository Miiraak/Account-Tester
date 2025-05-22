namespace AccountTester
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
            labelLogs = new Label();
            menuStrip1 = new MenuStrip();
            optionsToolStripMenuItem = new ToolStripMenuItem();
            langageToolStripMenuItem = new ToolStripMenuItem();
            enUSToolStripMenuItem = new ToolStripMenuItem();
            frFRToolStripMenuItem = new ToolStripMenuItem();
            label1 = new Label();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // buttonStart
            // 
            buttonStart.Location = new Point(12, 352);
            buttonStart.Name = "buttonStart";
            buttonStart.Size = new Size(88, 21);
            buttonStart.TabIndex = 0;
            buttonStart.Text = "Start";
            buttonStart.UseVisualStyleBackColor = true;
            buttonStart.Click += ButtonStart_Click;
            // 
            // richTextBoxLogs
            // 
            richTextBoxLogs.HideSelection = false;
            richTextBoxLogs.Location = new Point(12, 56);
            richTextBoxLogs.Name = "richTextBoxLogs";
            richTextBoxLogs.ReadOnly = true;
            richTextBoxLogs.Size = new Size(304, 290);
            richTextBoxLogs.TabIndex = 4;
            richTextBoxLogs.TabStop = false;
            richTextBoxLogs.Text = "";
            // 
            // buttonExportForm
            // 
            buttonExportForm.Location = new Point(239, 352);
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
            buttonCopier.Location = new Point(132, 352);
            buttonCopier.Name = "buttonCopier";
            buttonCopier.Size = new Size(75, 21);
            buttonCopier.TabIndex = 6;
            buttonCopier.Text = "Copy";
            buttonCopier.UseVisualStyleBackColor = true;
            buttonCopier.Click += ButtonCopier_Click;
            // 
            // labelLogs
            // 
            labelLogs.AutoSize = true;
            labelLogs.Font = new Font("Consolas", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            labelLogs.Location = new Point(12, 35);
            labelLogs.Name = "labelLogs";
            labelLogs.Size = new Size(56, 18);
            labelLogs.TabIndex = 3;
            labelLogs.Text = "Logs :";
            // 
            // menuStrip1
            // 
            menuStrip1.BackColor = SystemColors.ButtonFace;
            menuStrip1.Items.AddRange(new ToolStripItem[] { optionsToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(328, 24);
            menuStrip1.TabIndex = 7;
            menuStrip1.Text = "menuStrip1";
            // 
            // optionsToolStripMenuItem
            // 
            optionsToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { langageToolStripMenuItem });
            optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            optionsToolStripMenuItem.Size = new Size(61, 20);
            optionsToolStripMenuItem.Text = "Options";
            // 
            // langageToolStripMenuItem
            // 
            langageToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { enUSToolStripMenuItem, frFRToolStripMenuItem });
            langageToolStripMenuItem.Name = "langageToolStripMenuItem";
            langageToolStripMenuItem.Size = new Size(119, 22);
            langageToolStripMenuItem.Text = "Langage";
            // 
            // enUSToolStripMenuItem
            // 
            enUSToolStripMenuItem.Name = "enUSToolStripMenuItem";
            enUSToolStripMenuItem.Size = new Size(106, 22);
            enUSToolStripMenuItem.Text = "en-US";
            enUSToolStripMenuItem.Click += enUSToolStripMenuItem_Click;
            // 
            // frFRToolStripMenuItem
            // 
            frFRToolStripMenuItem.Name = "frFRToolStripMenuItem";
            frFRToolStripMenuItem.Size = new Size(106, 22);
            frFRToolStripMenuItem.Text = "fr-FR";
            frFRToolStripMenuItem.Click += frFRToolStripMenuItem_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(0, 12);
            label1.Name = "label1";
            label1.Size = new Size(329, 14);
            label1.TabIndex = 8;
            label1.Text = "______________________________________________";
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
            Controls.Add(labelLogs);
            Controls.Add(buttonStart);
            Controls.Add(menuStrip1);
            Controls.Add(label1);
            Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
            MaximizeBox = false;
            Name = "MainForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Account Tester v0.7.8";
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button buttonStart;
        private RichTextBox richTextBoxLogs;
        private Button buttonExportForm;
        private Button buttonCopier;
        private Label labelLogs;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem optionsToolStripMenuItem;
        private ToolStripMenuItem langageToolStripMenuItem;
        private ToolStripMenuItem enUSToolStripMenuItem;
        private ToolStripMenuItem frFRToolStripMenuItem;
        private Label label1;
    }
}
