using System.Globalization;
using System.Resources;

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

        public string CurrentCulture => _culture.Name;
    }
}
