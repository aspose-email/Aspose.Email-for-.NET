using System;
using System.Collections.Generic;
using System.Text;

namespace Aspose.Email.Live.Demos.UI.Parsing
{
    /// <summary>
    /// Universal Lexemme parser.
    /// Parses a string according to the rules of transitions between language tokens.
    /// It can parse any language if all the necessary rules are described.
    /// </summary>
    public class LexemmeParser
    {
        private ISwitchRule[] _rules;
        private Func<char, string> _charTypeSelector;

        public LexemmeParser(ISwitchRule[] rules, Func<char, string> charTypeSelector)
        {
            _rules = rules;
            _charTypeSelector = charTypeSelector;
        }

        public List<Lexemme> ParseString(string str)
        {
            string currentLexemmeType = null;
            string previousCharType = null;

            var list = new List<Lexemme>();
            var lexemmeBuilder = new StringBuilder();

            for (int i = 0; i < str.Length; i++)
            {
                var currentChar = char.ToLowerInvariant(str[i]);
                var type = _charTypeSelector(currentChar);

                for (int u = 0; u < _rules.Length; u++)
                {
                    var rule = _rules[u];

                    if ((rule.PreviousCharType == null || rule.PreviousCharType == previousCharType) && (rule.NextCharType == null || rule.NextCharType == type) && rule.CurrentLexemme == currentLexemmeType)
                    {
                        rule.UpdateLexemme(list, lexemmeBuilder, str[i], ref currentLexemmeType);
                        previousCharType = type;

                        goto CONTINUE;
                    }
                }

                lexemmeBuilder.Append(str[i]);
                previousCharType = type;

				CONTINUE: continue;
            }

            list.Add(new Lexemme(currentLexemmeType, lexemmeBuilder.ToString()));

            return list;
        }
    }
}
