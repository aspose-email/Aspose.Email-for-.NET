using System.Collections.Generic;
using System.Text;

namespace Aspose.Email.Live.Demos.UI.Parsing
{
    public class SimpleSwitchRule : ISwitchRule
    {
        public string PreviousCharType { get; }
        public string NextCharType { get; }

        public string CurrentLexemme { get; }
        public string Target { get; }

        public SimpleSwitchRule(string current, string next, string targetType)
        {
            CurrentLexemme = current;
            NextCharType = next;
            Target = targetType;
        }

        public SimpleSwitchRule(string current, string previous, string next, string targetType)
        {
            CurrentLexemme = current;
            PreviousCharType = previous;
            NextCharType = next;
            Target = targetType;
        }

        public void UpdateLexemme(List<Lexemme> lexemmes, StringBuilder builder, char? ch, ref string current)
        {
            if (Target != current)
            {
                if (current != null)
                {
                    lexemmes.Add(new Lexemme(current, builder.ToString()));
                    builder.Clear();
                }

                current = Target;
            }
            
            if (ch.HasValue)
                builder.Append(ch);
        }
    }
}
