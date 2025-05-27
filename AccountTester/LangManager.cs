using System.Globalization;
using System.Resources;
using System.Text;

namespace AccountTester
{
    internal class LangManager
    {
        private static LangManager? _instance;
        private readonly ResourceManager _resourceManager;
        private CultureInfo _culture;
        public event Action? LanguageChanged;

        /// <summary>
        /// Singleton instance of the LangManager.
        /// </summary>
        private LangManager()
        {
            _resourceManager = new ResourceManager("AccountTester.strings", typeof(LangManager).Assembly);
            _culture = new CultureInfo("en-US");
        }

        /// <summary>
        /// Gets the singleton instance of the LangManager.
        /// </summary>
        public static LangManager Instance
        {
            get
            {
                _instance ??= new LangManager();
                return _instance;
            }
        }

        /// <summary>
        /// Sets the language for translations.
        /// </summary>
        /// <param name="cultureCode"></param>
        public void SetLanguage(string cultureCode)
        {
            _culture = new CultureInfo(cultureCode);
            LanguageChanged?.Invoke();
        }

        /// <summary>
        /// Translates a given key to the current language.
        /// </summary>
        /// <param name="key">The name of the string in the resource file.</param>
        /// <returns>Returns the value from the resource file if found, otherwise returns the key wrapped in exclamation marks.</returns>
        public string Translate(string key)
        {
            return _resourceManager.GetString(key, _culture) ?? $"!{key}!";
        }

        /// <summary>
        /// Translates a given key to the current language, remove the diacritics and trims it to a single word in PascalCase format. 
        /// </summary>
        /// <param name="key">The name of the string in the resource file.</param>
        /// <returns>Returns the result as a single word without spaces.</returns>
        public string TrimTranslate(string key)
        {
            var translation = _resourceManager.GetString(key, _culture) ?? $"!{key}!";
            var words = translation
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(word => RemoveDiacritics(word))
                .Select(word => char.ToUpper(word[0]) + word.Substring(1).ToLowerInvariant());

            return string.Join("", words);
        }

        /// <summary>
        /// Removes diacritics from a string.
        /// </summary>
        static string RemoveDiacritics(string text)
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
    }
}
