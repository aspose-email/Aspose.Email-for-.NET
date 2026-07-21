// Demonstrates a Windows Forms UI for validating an email address
// using the Aspose.Email EmailValidator.

using Aspose.Email.Tools.Verifications;
using System;
using System.Windows.Forms;

namespace Aspose.Email.Examples.Email
{
    public partial class EmailValidations : Form
    {
        public EmailValidations()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var ev = new EmailValidator();
            try
            {
                ev.Validate(txtEmailAddr.Text, out var result);
                lblResult.Text = result.ReturnCode == ValidationResponseCode.ValidationSuccess
                    ? "The email address is valid."
                    : $"The email address is invalid. Return code: {result.ReturnCode}.";
            }
            catch (Exception ex)
            {
                lblResult.Text = ex.Message;
            }
        }
    }
}
