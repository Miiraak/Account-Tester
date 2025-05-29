namespace AccountTester
{
    partial class ContactForm
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
            linkLabelMail = new LinkLabel();
            labelMail = new Label();
            labelRepository = new Label();
            linkLabelRepository = new LinkLabel();
            labelSite = new Label();
            linkLabelSite = new LinkLabel();
            SuspendLayout();
            // 
            // linkLabelMail
            // 
            linkLabelMail.AutoSize = true;
            linkLabelMail.Location = new Point(114, 28);
            linkLabelMail.Name = "linkLabelMail";
            linkLabelMail.Size = new Size(133, 14);
            linkLabelMail.TabIndex = 0;
            linkLabelMail.TabStop = true;
            linkLabelMail.Tag = "A";
            linkLabelMail.Text = "miiraak@miiraak.ch";
            linkLabelMail.LinkClicked += OpenContactInfo;
            // 
            // labelMail
            // 
            labelMail.AutoSize = true;
            labelMail.Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            labelMail.Location = new Point(17, 28);
            labelMail.Name = "labelMail";
            labelMail.Size = new Size(49, 14);
            labelMail.TabIndex = 1;
            labelMail.Text = "Mail :";
            // 
            // labelRepository
            // 
            labelRepository.AutoSize = true;
            labelRepository.Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            labelRepository.Location = new Point(17, 66);
            labelRepository.Name = "labelRepository";
            labelRepository.Size = new Size(91, 14);
            labelRepository.TabIndex = 3;
            labelRepository.Text = "Repository :";
            // 
            // linkLabelRepository
            // 
            linkLabelRepository.AutoSize = true;
            linkLabelRepository.Location = new Point(114, 66);
            linkLabelRepository.Name = "linkLabelRepository";
            linkLabelRepository.Size = new Size(161, 14);
            linkLabelRepository.TabIndex = 2;
            linkLabelRepository.TabStop = true;
            linkLabelRepository.Tag = "B";
            linkLabelRepository.Text = "Miiraak/Account-Tester";
            linkLabelRepository.LinkClicked += OpenContactInfo;
            // 
            // labelSite
            // 
            labelSite.AutoSize = true;
            labelSite.Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            labelSite.Location = new Point(17, 105);
            labelSite.Name = "labelSite";
            labelSite.Size = new Size(49, 14);
            labelSite.TabIndex = 5;
            labelSite.Text = "Site :";
            // 
            // linkLabelSite
            // 
            linkLabelSite.AutoSize = true;
            linkLabelSite.Location = new Point(114, 105);
            linkLabelSite.Name = "linkLabelSite";
            linkLabelSite.Size = new Size(105, 14);
            linkLabelSite.TabIndex = 4;
            linkLabelSite.TabStop = true;
            linkLabelSite.Tag = "C";
            linkLabelSite.Text = "www.miiraak.ch";
            linkLabelSite.LinkClicked += OpenContactInfo;
            // 
            // ContactForm
            // 
            AutoScaleDimensions = new SizeF(7F, 14F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            ClientSize = new Size(284, 150);
            Controls.Add(labelSite);
            Controls.Add(linkLabelSite);
            Controls.Add(labelRepository);
            Controls.Add(linkLabelRepository);
            Controls.Add(labelMail);
            Controls.Add(linkLabelMail);
            Font = new Font("Consolas", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "ContactForm";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Contact";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private LinkLabel linkLabelMail;
        private Label labelMail;
        private Label labelRepository;
        private LinkLabel linkLabelRepository;
        private Label labelSite;
        private LinkLabel linkLabelSite;
    }
}