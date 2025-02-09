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
            textBoxFileName = new TextBox();
            textBoxFilePath = new TextBox();
            comboBoxExtension = new ComboBox();
            label1 = new Label();
            label2 = new Label();
            label3 = new Label();
            buttonExport = new Button();
            buttonSelectPath = new Button();
            SuspendLayout();
            // 
            // textBoxFileName
            // 
            textBoxFileName.Location = new Point(12, 25);
            textBoxFileName.Name = "textBoxFileName";
            textBoxFileName.Size = new Size(188, 22);
            textBoxFileName.TabIndex = 1;
            // 
            // textBoxFilePath
            // 
            textBoxFilePath.Location = new Point(10, 76);
            textBoxFilePath.Name = "textBoxFilePath";
            textBoxFilePath.ReadOnly = true;
            textBoxFilePath.Size = new Size(136, 22);
            textBoxFilePath.TabIndex = 2;
            // 
            // comboBoxExtension
            // 
            comboBoxExtension.FormattingEnabled = true;
            comboBoxExtension.Items.AddRange(new object[] { ".txt", ".pdf", ".csv", ".xml", ".json", ".xlsx" });
            comboBoxExtension.Location = new Point(10, 127);
            comboBoxExtension.Name = "comboBoxExtension";
            comboBoxExtension.Size = new Size(188, 22);
            comboBoxExtension.TabIndex = 4;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(10, 110);
            label1.Name = "label1";
            label1.Size = new Size(84, 14);
            label1.TabIndex = 5;
            label1.Text = "Extension :";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(12, 8);
            label2.Name = "label2";
            label2.Size = new Size(42, 14);
            label2.TabIndex = 6;
            label2.Text = "Nom :";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(10, 59);
            label3.Name = "label3";
            label3.Size = new Size(63, 14);
            label3.TabIndex = 7;
            label3.Text = "Chemin :";
            // 
            // buttonExport
            // 
            buttonExport.Location = new Point(71, 155);
            buttonExport.Name = "buttonExport";
            buttonExport.Size = new Size(75, 23);
            buttonExport.TabIndex = 8;
            buttonExport.Text = "Export";
            buttonExport.UseVisualStyleBackColor = true;
            buttonExport.Click += ButtonExport_Click;
            // 
            // buttonSelectPath
            // 
            buttonSelectPath.Location = new Point(152, 76);
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
            ClientSize = new Size(216, 189);
            Controls.Add(buttonSelectPath);
            Controls.Add(buttonExport);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(comboBoxExtension);
            Controls.Add(textBoxFilePath);
            Controls.Add(textBoxFileName);
            Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            Name = "ExportForm";
            Text = "Exportation";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private TextBox textBoxFileName;
        private TextBox textBoxFilePath;
        private ComboBox comboBoxExtension;
        private Label label1;
        private Label label2;
        private Label label3;
        private Button buttonExport;
        private Button buttonSelectPath;
    }
}