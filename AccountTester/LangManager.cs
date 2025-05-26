using System.Globalization;
using System.Resources;
using System.Text;

namespace AccountTester
{
    internal class LangManager
    {
        private static LangManager _instance;
        private ResourceManager _resourceManager;
        private CultureInfo _culture;
        public event Action? LanguageChanged;

        private LangManager()
        {
            _resourceManager = new ResourceManager("AccountTester.strings", typeof(LangManager).Assembly);
            _culture = new CultureInfo("en-US");
        }

        public static LangManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new LangManager();
                }
                return _instance;
            }
        }

        public void SetLanguage(string cultureCode)
        {
            _culture = new CultureInfo(cultureCode);
            LanguageChanged?.Invoke();
        }

        public string Translate(string key)
        {
            return _resourceManager.GetString(key, _culture) ?? $"!{key}!";
        }

        public string TrimTranslate(string key)
        {
            var translation = _resourceManager.GetString(key, _culture) ?? $"!{key}!";
            var words = translation
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(word => RemoveDiacritics(word))
                .Select(word => char.ToUpper(word[0]) + word.Substring(1).ToLowerInvariant());

            return string.Join("", words);
        }

        string RemoveDiacritics(string text)
        {
            var normalized = text.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();

            foreach (var c in normalized)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                    sb.Append(c);
            }

            return sb.ToString().Normalize(NormalizationForm.FormC);
        }

        public string CurrentCulture => _culture.Name;
    }
}
