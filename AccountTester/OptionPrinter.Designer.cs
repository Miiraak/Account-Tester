namespace AccountTester
{
    partial class OptionPrinter
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
            labelPrinterList = new Label();
            textBoxPrinterAdd = new TextBox();
            buttonAdd = new Button();
            checkedListBoxPrinter = new CheckedListBox();
            SuspendLayout();
            // 
            // labelPrinterList
            // 
            labelPrinterList.AutoSize = true;
            labelPrinterList.Location = new Point(12, 9);
            labelPrinterList.Name = "labelPrinterList";
            labelPrinterList.Size = new Size(105, 14);
            labelPrinterList.TabIndex = 0;
            labelPrinterList.Text = "Printer List :";
            // 
            // textBoxPrinterAdd
            // 
            textBoxPrinterAdd.Location = new Point(12, 26);
            textBoxPrinterAdd.Name = "textBoxPrinterAdd";
            textBoxPrinterAdd.Size = new Size(184, 22);
            textBoxPrinterAdd.TabIndex = 1;
            // 
            // buttonAdd
            // 
            buttonAdd.Location = new Point(202, 26);
            buttonAdd.Name = "buttonAdd";
            buttonAdd.Size = new Size(65, 21);
            buttonAdd.TabIndex = 2;
            buttonAdd.Text = "Add";
            buttonAdd.UseVisualStyleBackColor = true;
            buttonAdd.Click += ButtonAdd_Click;
            // 
            // checkedListBoxPrinter
            // 
            checkedListBoxPrinter.CheckOnClick = true;
            checkedListBoxPrinter.FormattingEnabled = true;
            checkedListBoxPrinter.Location = new Point(12, 54);
            checkedListBoxPrinter.Name = "checkedListBoxPrinter";
            checkedListBoxPrinter.Size = new Size(255, 174);
            checkedListBoxPrinter.TabIndex = 3;
            checkedListBoxPrinter.SelectedIndexChanged += RemovePrinter;
            // 
            // OptionPrinter
            // 
            AutoScaleDimensions = new SizeF(7F, 14F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            ClientSize = new Size(281, 245);
            Controls.Add(checkedListBoxPrinter);
            Controls.Add(buttonAdd);
            Controls.Add(textBoxPrinterAdd);
            Controls.Add(labelPrinterList);
            Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "OptionPrinter";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Printer options";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label labelPrinterList;
        private TextBox textBoxPrinterAdd;
        private Button buttonAdd;
        private CheckedListBox checkedListBoxPrinter;
    }
}