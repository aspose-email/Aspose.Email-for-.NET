namespace Aspose.Email.Live.Demos.UI.Parsing.Html
{
    public static class HtmlLexemmeParserRules
    {
        /// <summary>
        /// Rules for parsing HTML text via LexemmeParser
        /// </summary>
        public static ISwitchRule[] DefaultRules => new[]
        {
             new SimpleSwitchRule(null,                                   nameof(CharType.LeftAngle),    nameof(HtmlLexemmeType.OpenCorner)),
             new SimpleSwitchRule(null,                                   nameof(CharType.Space),        nameof(HtmlLexemmeType.Text)),
             new SimpleSwitchRule(null,                                   nameof(CharType.Unknown),      nameof(HtmlLexemmeType.Text)),

             new SimpleSwitchRule(nameof(HtmlLexemmeType.OpenCorner),     nameof(CharType.Exclamation),  nameof(HtmlLexemmeType.Comment)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.OpenCorner),     nameof(CharType.Unknown),      nameof(HtmlLexemmeType.Element)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.OpenCorner),     nameof(CharType.Space),        nameof(HtmlLexemmeType.Space)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.OpenCorner),     nameof(CharType.RightSlash),   nameof(HtmlLexemmeType.RightSlash)),

             new SimpleSwitchRule(nameof(HtmlLexemmeType.Comment),        nameof(CharType.Dash),		 nameof(CharType.RightAngle),			nameof(HtmlLexemmeType.CloseCorner)),

             new SimpleSwitchRule(nameof(HtmlLexemmeType.RightSlash),     nameof(CharType.RightAngle),   nameof(HtmlLexemmeType.CloseCorner)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.RightSlash),     null,                          nameof(HtmlLexemmeType.Element)),

             new SimpleSwitchRule(nameof(HtmlLexemmeType.Element),        nameof(CharType.Space),        nameof(HtmlLexemmeType.Space)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.Element),        nameof(CharType.Equally),      nameof(HtmlLexemmeType.EqualSign)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.Element),        nameof(CharType.RightAngle),   nameof(HtmlLexemmeType.CloseCorner)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.Element),        nameof(CharType.RightSlash),   nameof(HtmlLexemmeType.RightSlash)),

             new SimpleSwitchRule(nameof(HtmlLexemmeType.EqualSign),      nameof(CharType.Quotes),       nameof(HtmlLexemmeType.OpenQuotes)),

             new SimpleSwitchRule(nameof(HtmlLexemmeType.Space),          nameof(CharType.RightSlash),   nameof(HtmlLexemmeType.RightSlash)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.Space),          nameof(CharType.RightAngle),   nameof(HtmlLexemmeType.CloseCorner)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.Space),          nameof(CharType.Unknown),      nameof(HtmlLexemmeType.AttributeName)),

             new SimpleSwitchRule(nameof(HtmlLexemmeType.AttributeName),  nameof(CharType.Equally),      nameof(HtmlLexemmeType.EqualSign)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.AttributeName),  nameof(CharType.RightSlash),      nameof(HtmlLexemmeType.RightSlash)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.AttributeName),  nameof(CharType.Quotes),      nameof(HtmlLexemmeType.OpenQuotes)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.AttributeName),  nameof(CharType.Space),      nameof(HtmlLexemmeType.Space)),

             new SimpleSwitchRule(nameof(HtmlLexemmeType.OpenQuotes),     nameof(CharType.Quotes),       nameof(HtmlLexemmeType.CloseQuotes)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.OpenQuotes),     null,                          nameof(HtmlLexemmeType.AttributeValue)),

             new SimpleSwitchRule(nameof(HtmlLexemmeType.AttributeValue), nameof(CharType.Quotes),       nameof(HtmlLexemmeType.CloseQuotes)),

             new SimpleSwitchRule(nameof(HtmlLexemmeType.CloseQuotes),    nameof(CharType.RightAngle),   nameof(HtmlLexemmeType.CloseCorner)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.CloseQuotes),    nameof(CharType.Space),        nameof(HtmlLexemmeType.Space)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.CloseQuotes),    nameof(CharType.RightSlash),   nameof(HtmlLexemmeType.RightSlash)),

             new SimpleSwitchRule(nameof(HtmlLexemmeType.Text),           nameof(CharType.LeftAngle),    nameof(HtmlLexemmeType.OpenCorner)),

             new SimpleSwitchRule(nameof(HtmlLexemmeType.CloseCorner),    nameof(CharType.LeftAngle),    nameof(HtmlLexemmeType.OpenCorner)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.CloseCorner),    nameof(CharType.Unknown),      nameof(HtmlLexemmeType.Text)),
             new SimpleSwitchRule(nameof(HtmlLexemmeType.CloseCorner),    nameof(CharType.Space),        nameof(HtmlLexemmeType.Text))
        };

        public static string DefaultCharTypeSelector(char ch)
        {
            switch (ch)
            {
                case ' ': return nameof(CharType.Space);
                case '<': return nameof(CharType.LeftAngle);
                case '>': return nameof(CharType.RightAngle);
                case '"': return nameof(CharType.Quotes);
                case '/': return nameof(CharType.RightSlash);
                case '=': return nameof(CharType.Equally);
                case '-': return nameof(CharType.Dash);
                default: return nameof(CharType.Unknown);
            }
        }
    }
}
