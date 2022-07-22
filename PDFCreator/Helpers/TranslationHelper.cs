using PDFCreator.Enums;
using System;
using System.Collections.Generic;
using System.Resources;
using System.Text;
using System.Xml;

namespace PDFCreator.Helpers
{
    public static class TranslationHelper
    {
        #region "Declarations"
        // downloaded Xml document
        private static XmlDocument xmlDoc;
        // dictionary for the selected language
        private static Dictionary<string, string> _dictionary;
        // selected language
        private static Language language;
        #endregion

        #region "Constructor"
        /// <summary>
        /// Download Xml document and fill dictionary for the bulgarian language
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        static TranslationHelper()
        {
            xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(Resources.Resources.Dictionary);

            _dictionary = new Dictionary<string, string>();
            Language = Language.BG;            
        }
        #endregion

        #region "Properties"
        /// <summary>
        /// Selected language
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        public static Language Language
        {
            get => language;
            set
            {
                language = value;

                FillDictionary();
            }
        }
        #endregion

        #region "Methods"
        /// <summary>
        /// Get localized text
        /// </summary>
        /// <param name="key">Dictionary search key</param>
        /// <returns>Localized text</returns>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        public static string GetLocalizedString(string key)
        {
            if (_dictionary.ContainsKey(key))
            {
                return _dictionary[key];
            }
            else
            {
                return "NOT TRANSLATED";
            }
        }

        /// <summary>
        /// Fill dictionary in according to selected language
        /// </summary>
        /// <developer>Serhii Rozniuk</developer>
        /// <date>11.03.2021</date>
        private static void FillDictionary()
        {
            XmlElement root = xmlDoc.DocumentElement;
            string key, value;

            if (_dictionary.Count > 0)
            {
                _dictionary.Clear();
            }

            foreach (XmlNode node in root)
            {
                key = string.Empty;
                value = string.Empty;
                if (node.Name.Equals("Data"))
                {
                    if (node.Attributes.Count > 0)
                    {
                        key = node.Attributes[0].Value;
                    }
                    foreach (XmlNode childNode in node.ChildNodes)
                    {
                        if (childNode.Attributes.Count > 0 && childNode.Attributes[0].Value.Equals(Language.ToString()))
                        {
                            value = childNode.InnerText;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
                {
                    _dictionary.Add(key, value);
                }
            }
        }
        #endregion
    }
}
