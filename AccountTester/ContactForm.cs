namespace AccountTester
{
    public partial class ContactForm : Form
    {
        static string T(string key) => LangManager.Instance.Translate(key);

        public ContactForm()
        {
            InitializeComponent();

            UpdateTexts();
            LangManager.Instance.LanguageChanged += UpdateTexts;
        }

        private void UpdateTexts()
        {
            this.Text = T("Contact");
            labelMail.Text = $"{T("Mail")} :";
            labelRepository.Text = $"{T("Repository")} :";
            labelSite.Text = $"{T("Site")} :";
        }

        private void OpenContactInfo(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel link = (LinkLabel)sender;
            string source = String.Empty;

            switch (link.Tag)
            {
                case "A":
                    source = "mailto:miiraak@miiraak.ch";
                    break;
                case "B":
                    source = "https://github.com/Miiraak/Account-Tester";
                    break;
                case "C":
                    source = "https://www.miiraak.ch";
                    break;
            }

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(source) { UseShellExecute = true });
        }
    }
}
