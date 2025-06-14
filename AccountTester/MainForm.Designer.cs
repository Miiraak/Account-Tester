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
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            richTextBoxLogs = new RichTextBox();
            contextMenuStrip1 = new ContextMenuStrip(components);
            copyToolStripMenuItem = new ToolStripMenuItem();
            labelLogs = new Label();
            menuStrip1 = new MenuStrip();
            startToolStripMenuItem = new ToolStripMenuItem();
            exportToolStripMenuItem = new ToolStripMenuItem();
            optionsToolStripMenuItem = new ToolStripMenuItem();
            autorunToolStripMenuItem = new ToolStripMenuItem();
            languageToolStripMenuItem = new ToolStripMenuItem();
            enUSToolStripMenuItem = new ToolStripMenuItem();
            frFRToolStripMenuItem = new ToolStripMenuItem();
            reportsToolStripMenuItem = new ToolStripMenuItem();
            autoExportToolStripMenuItem = new ToolStripMenuItem();
            extensionByDefaultToolStripMenuItem = new ToolStripMenuItem();
            toolStripComboBoxExtensionByDefault = new ToolStripComboBox();
            toolStripSeparator4 = new ToolStripSeparator();
            saveToolStripMenuItem = new ToolStripMenuItem();
            helpToolStripMenuItem = new ToolStripMenuItem();
            contactToolStripMenuItem = new ToolStripMenuItem();
            label1 = new Label();
            contextMenuStrip1.SuspendLayout();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // richTextBoxLogs
            // 
            richTextBoxLogs.ContextMenuStrip = contextMenuStrip1;
            richTextBoxLogs.Cursor = Cursors.Cross;
            richTextBoxLogs.HideSelection = false;
            richTextBoxLogs.Location = new Point(12, 56);
            richTextBoxLogs.Name = "richTextBoxLogs";
            richTextBoxLogs.ReadOnly = true;
            richTextBoxLogs.Size = new Size(304, 313);
            richTextBoxLogs.TabIndex = 4;
            richTextBoxLogs.TabStop = false;
            richTextBoxLogs.Text = "";
            // 
            // contextMenuStrip1
            // 
            contextMenuStrip1.Items.AddRange(new ToolStripItem[] { copyToolStripMenuItem });
            contextMenuStrip1.Name = "contextMenuStrip1";
            contextMenuStrip1.Size = new Size(103, 26);
            // 
            // copyToolStripMenuItem
            // 
            copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            copyToolStripMenuItem.Size = new Size(102, 22);
            copyToolStripMenuItem.Text = "Copy";
            copyToolStripMenuItem.Click += CopyToolStripMenuItem_Click;
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
            menuStrip1.Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            menuStrip1.Items.AddRange(new ToolStripItem[] { startToolStripMenuItem, exportToolStripMenuItem, optionsToolStripMenuItem, helpToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(328, 24);
            menuStrip1.TabIndex = 7;
            menuStrip1.Text = "menuStrip1";
            // 
            // startToolStripMenuItem
            // 
            startToolStripMenuItem.Name = "startToolStripMenuItem";
            startToolStripMenuItem.Size = new Size(54, 20);
            startToolStripMenuItem.Text = "Start";
            startToolStripMenuItem.Click += StartToolStripMenuItem_Click;
            // 
            // exportToolStripMenuItem
            // 
            exportToolStripMenuItem.Name = "exportToolStripMenuItem";
            exportToolStripMenuItem.Size = new Size(61, 20);
            exportToolStripMenuItem.Text = "Export";
            exportToolStripMenuItem.Click += ExportToolStripMenuItem_Click;
            // 
            // optionsToolStripMenuItem
            // 
            optionsToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { autorunToolStripMenuItem, languageToolStripMenuItem, reportsToolStripMenuItem, toolStripSeparator4, saveToolStripMenuItem });
            optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            optionsToolStripMenuItem.Size = new Size(68, 20);
            optionsToolStripMenuItem.Text = "Options";
            // 
            // autorunToolStripMenuItem
            // 
            autorunToolStripMenuItem.Name = "autorunToolStripMenuItem";
            autorunToolStripMenuItem.Size = new Size(130, 22);
            autorunToolStripMenuItem.Text = "Autorun";
            autorunToolStripMenuItem.Click += AutorunToolStripMenuItem_Click;
            // 
            // languageToolStripMenuItem
            // 
            languageToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { enUSToolStripMenuItem, frFRToolStripMenuItem });
            languageToolStripMenuItem.Name = "languageToolStripMenuItem";
            languageToolStripMenuItem.Size = new Size(130, 22);
            languageToolStripMenuItem.Text = "Language";
            // 
            // enUSToolStripMenuItem
            // 
            enUSToolStripMenuItem.Checked = true;
            enUSToolStripMenuItem.CheckState = CheckState.Checked;
            enUSToolStripMenuItem.Name = "enUSToolStripMenuItem";
            enUSToolStripMenuItem.Size = new Size(109, 22);
            enUSToolStripMenuItem.Tag = "";
            enUSToolStripMenuItem.Text = "en-US";
            enUSToolStripMenuItem.Click += EnUSToolStripMenuItem_Click;
            // 
            // frFRToolStripMenuItem
            // 
            frFRToolStripMenuItem.Name = "frFRToolStripMenuItem";
            frFRToolStripMenuItem.Size = new Size(109, 22);
            frFRToolStripMenuItem.Tag = "";
            frFRToolStripMenuItem.Text = "fr-FR";
            frFRToolStripMenuItem.Click += FrFRToolStripMenuItem_Click;
            // 
            // reportsToolStripMenuItem
            // 
            reportsToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { autoExportToolStripMenuItem, extensionByDefaultToolStripMenuItem, toolStripComboBoxExtensionByDefault });
            reportsToolStripMenuItem.Name = "reportsToolStripMenuItem";
            reportsToolStripMenuItem.Size = new Size(130, 22);
            reportsToolStripMenuItem.Text = "Reports";
            // 
            // autoExportToolStripMenuItem
            // 
            autoExportToolStripMenuItem.CheckOnClick = true;
            autoExportToolStripMenuItem.Name = "autoExportToolStripMenuItem";
            autoExportToolStripMenuItem.Size = new Size(228, 22);
            autoExportToolStripMenuItem.Text = "Auto-Export";
            // 
            // extensionByDefaultToolStripMenuItem
            // 
            extensionByDefaultToolStripMenuItem.Name = "extensionByDefaultToolStripMenuItem";
            extensionByDefaultToolStripMenuItem.Size = new Size(228, 22);
            extensionByDefaultToolStripMenuItem.Text = "Extension by default :";
            // 
            // toolStripComboBoxExtensionByDefault
            // 
            toolStripComboBoxExtensionByDefault.Items.AddRange(new object[] { ".log", ".txt", ".csv", ".xml", ".json", ".zip" });
            toolStripComboBoxExtensionByDefault.Name = "toolStripComboBoxExtensionByDefault";
            toolStripComboBoxExtensionByDefault.Size = new Size(121, 23);
            // 
            // toolStripSeparator4
            // 
            toolStripSeparator4.Name = "toolStripSeparator4";
            toolStripSeparator4.Size = new Size(127, 6);
            // 
            // saveToolStripMenuItem
            // 
            saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            saveToolStripMenuItem.Size = new Size(130, 22);
            saveToolStripMenuItem.Text = "Save";
            saveToolStripMenuItem.Click += SaveToolStripMenuItem_Click;
            // 
            // helpToolStripMenuItem
            // 
            helpToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { contactToolStripMenuItem });
            helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            helpToolStripMenuItem.Size = new Size(47, 20);
            helpToolStripMenuItem.Text = "Help";
            // 
            // contactToolStripMenuItem
            // 
            contactToolStripMenuItem.Name = "contactToolStripMenuItem";
            contactToolStripMenuItem.Size = new Size(123, 22);
            contactToolStripMenuItem.Text = "Contact";
            contactToolStripMenuItem.Click += ContactToolStripMenuItem_Click;
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
            Controls.Add(richTextBoxLogs);
            Controls.Add(labelLogs);
            Controls.Add(menuStrip1);
            Controls.Add(label1);
            Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = menuStrip1;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "MainForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Account Tester v0.8";
            Load += MainFormLoad;
            contextMenuStrip1.ResumeLayout(false);
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private RichTextBox richTextBoxLogs;
        private Label labelLogs;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem optionsToolStripMenuItem;
        private ToolStripMenuItem languageToolStripMenuItem;
        private ToolStripMenuItem enUSToolStripMenuItem;
        private ToolStripMenuItem frFRToolStripMenuItem;
        private Label label1;
        private ToolStripMenuItem helpToolStripMenuItem;
        private ToolStripMenuItem contactToolStripMenuItem;
        private ContextMenuStrip contextMenuStrip1;
        private ToolStripMenuItem copyToolStripMenuItem;
        private ToolStripMenuItem exportToolStripMenuItem;
        private ToolStripMenuItem startToolStripMenuItem;
        private ToolStripMenuItem reportsToolStripMenuItem;
        private ToolStripMenuItem autoExportToolStripMenuItem;
        private ToolStripMenuItem extensionByDefaultToolStripMenuItem;
        private ToolStripComboBox toolStripComboBoxExtensionByDefault;
        private ToolStripSeparator toolStripSeparator4;
        private ToolStripMenuItem saveToolStripMenuItem;
        private ToolStripMenuItem autorunToolStripMenuItem;
    }
}
