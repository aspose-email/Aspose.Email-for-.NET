using System.Collections.Generic;
using System.Text;

namespace Aspose.Email.Live.Demos.UI.Parsing
{
    /// <summary>
    /// Switching rules determine the transition condition between tokens when parsing a string.
    /// </summary>
    public interface ISwitchRule
    {
        string PreviousCharType { get; }
        string NextCharType { get; }
        string CurrentLexemme { get; }
        string Target { get; }

        void UpdateLexemme(List<Lexemme> lexemmes, StringBuilder builder, char? ch, ref string currentType);
    }
}
