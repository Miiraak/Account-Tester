namespace AccountTester
{
    partial class ExportForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExportForm));
            textBoxFileName = new TextBox();
            textBoxFilePath = new TextBox();
            comboBoxExtension = new ComboBox();
            labelExtension = new Label();
            labelFile = new Label();
            labelPath = new Label();
            buttonExport = new Button();
            buttonSelectPath = new Button();
            SuspendLayout();
            // 
            // textBoxFileName
            // 
            textBoxFileName.Location = new Point(12, 25);
            textBoxFileName.Name = "textBoxFileName";
            textBoxFileName.Size = new Size(221, 22);
            textBoxFileName.TabIndex = 1;
            // 
            // textBoxFilePath
            // 
            textBoxFilePath.Location = new Point(10, 76);
            textBoxFilePath.Name = "textBoxFilePath";
            textBoxFilePath.ReadOnly = true;
            textBoxFilePath.Size = new Size(169, 22);
            textBoxFilePath.TabIndex = 2;
            // 
            // comboBoxExtension
            // 
            comboBoxExtension.FormattingEnabled = true;
            comboBoxExtension.Items.AddRange(new object[] { ".zip", ".log", ".txt", ".csv", ".xml", ".json" });
            comboBoxExtension.Location = new Point(10, 127);
            comboBoxExtension.Name = "comboBoxExtension";
            comboBoxExtension.Size = new Size(223, 22);
            comboBoxExtension.TabIndex = 4;
            // 
            // labelExtension
            // 
            labelExtension.AutoSize = true;
            labelExtension.Location = new Point(10, 110);
            labelExtension.Name = "labelExtension";
            labelExtension.Size = new Size(84, 14);
            labelExtension.TabIndex = 5;
            labelExtension.Text = "Extension :";
            // 
            // labelFile
            // 
            labelFile.AutoSize = true;
            labelFile.Location = new Point(12, 8);
            labelFile.Name = "labelFile";
            labelFile.Size = new Size(84, 14);
            labelFile.TabIndex = 6;
            labelFile.Text = "File name :";
            // 
            // labelPath
            // 
            labelPath.AutoSize = true;
            labelPath.Location = new Point(10, 59);
            labelPath.Name = "labelPath";
            labelPath.Size = new Size(49, 14);
            labelPath.TabIndex = 7;
            labelPath.Text = "Path :";
            // 
            // buttonExport
            // 
            buttonExport.Location = new Point(85, 157);
            buttonExport.Name = "buttonExport";
            buttonExport.Size = new Size(75, 23);
            buttonExport.TabIndex = 8;
            buttonExport.Text = "Export";
            buttonExport.UseVisualStyleBackColor = true;
            buttonExport.Click += ButtonExport_Click;
            // 
            // buttonSelectPath
            // 
            buttonSelectPath.Location = new Point(185, 76);
            buttonSelectPath.Name = "buttonSelectPath";
            buttonSelectPath.Size = new Size(48, 23);
            buttonSelectPath.TabIndex = 9;
            buttonSelectPath.Text = "...";
            buttonSelectPath.UseVisualStyleBackColor = true;
            buttonSelectPath.Click += ButtonSelectPath_Click;
            // 
            // ExportForm
            // 
            AutoScaleDimensions = new SizeF(7F, 14F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            ClientSize = new Size(245, 189);
            Controls.Add(buttonSelectPath);
            Controls.Add(buttonExport);
            Controls.Add(labelPath);
            Controls.Add(labelFile);
            Controls.Add(labelExtension);
            Controls.Add(comboBoxExtension);
            Controls.Add(textBoxFilePath);
            Controls.Add(textBoxFileName);
            Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "ExportForm";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Export";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private TextBox textBoxFileName;
        private TextBox textBoxFilePath;
        private ComboBox comboBoxExtension;
        private Label labelExtension;
        private Label labelFile;
        private Label labelPath;
        private Button buttonExport;
        private Button buttonSelectPath;
    }
}