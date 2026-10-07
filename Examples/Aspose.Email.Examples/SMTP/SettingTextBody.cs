// Demonstrates setting a text body that is not plain ASCII.
//
// BodyEncoding decides how the text is encoded on the wire; UTF-8 carries any language,
// so accented letters, Cyrillic or CJK text arrive intact. SendPlainTextEmailMessage
// shows the ASCII-only case.

using System;
using System.Text;

namespace Aspose.Email.Examples.SMTP
{
    internal static class SettingTextBody
    {
        public static void Run()
        {
            if (!ClientBuilder.IsSmtpConfigured)
            {
                SmtpExampleInfo.PrintNotConfigured();
                return;
            }

            using (var client = ClientBuilder.Smtp(AuthType.Basic))
            {
                var message = new MailMessage(client.Username, client.Username)
                {
                    Subject = "Text body in UTF-8",
                    BodyEncoding = Encoding.UTF8,
                    Body = "Caf\u00e9, \u041f\u0440\u0438\u0432\u0435\u0442, \u3053\u3093\u306b\u3061\u306f"
                };

                client.Send(message);
                Console.WriteLine($"Sent a UTF-8 text body to {client.Username}.");
            }
        }
    }
}
