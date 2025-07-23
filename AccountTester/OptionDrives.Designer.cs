namespace AccountTester
{
    partial class OptionDrives
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
            labelDrivesList = new Label();
            ButtonAdd = new Button();
            checkedListBoxDrives = new CheckedListBox();
            textBoxDrivesAdd = new TextBox();
            SuspendLayout();
            // 
            // labelDrivesList
            // 
            labelDrivesList.AutoSize = true;
            labelDrivesList.Location = new Point(12, 9);
            labelDrivesList.Name = "labelDrivesList";
            labelDrivesList.Size = new Size(98, 14);
            labelDrivesList.TabIndex = 0;
            labelDrivesList.Text = "Drives List :";
            // 
            // ButtonAdd
            // 
            ButtonAdd.Location = new Point(202, 26);
            ButtonAdd.Name = "ButtonAdd";
            ButtonAdd.Size = new Size(65, 21);
            ButtonAdd.TabIndex = 3;
            ButtonAdd.Text = "Add";
            ButtonAdd.UseVisualStyleBackColor = true;
            ButtonAdd.Click += ButtonAdd_Click;
            // 
            // checkedListBoxDrives
            // 
            checkedListBoxDrives.CheckOnClick = true;
            checkedListBoxDrives.FormattingEnabled = true;
            checkedListBoxDrives.Location = new Point(12, 54);
            checkedListBoxDrives.Name = "checkedListBoxDrives";
            checkedListBoxDrives.Size = new Size(255, 174);
            checkedListBoxDrives.TabIndex = 5;
            checkedListBoxDrives.TabStop = false;
            checkedListBoxDrives.SelectedIndexChanged += RemoveDrives;
            // 
            // textBoxDrivesAdd
            // 
            textBoxDrivesAdd.Location = new Point(12, 26);
            textBoxDrivesAdd.Name = "textBoxDrivesAdd";
            textBoxDrivesAdd.Size = new Size(184, 22);
            textBoxDrivesAdd.TabIndex = 1;
            // 
            // OptionDrives
            // 
            AutoScaleDimensions = new SizeF(7F, 14F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            ClientSize = new Size(281, 245);
            Controls.Add(checkedListBoxDrives);
            Controls.Add(textBoxDrivesAdd);
            Controls.Add(ButtonAdd);
            Controls.Add(labelDrivesList);
            Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "OptionDrives";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Drives options";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label labelDrivesList;
        private Button ButtonAdd;
        private CheckedListBox checkedListBoxDrives;
        private TextBox textBoxDrivesAdd;
    }
}