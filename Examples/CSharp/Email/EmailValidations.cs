using Aspose.Email.Tools.Verifications;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET API reference 
when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Email for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Email.Examples.CSharp.Email
{
    public partial class EmailValidations : Form
    {
        public EmailValidations()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // ExStart:EmailValidations
            EmailValidator ev = new EmailValidator();
            ValidationResult result;
            try
            {
                ev.Validate(txtEmailAddr.Text, out result);
                if (result.ReturnCode == ValidationResponseCode.ValidationSuccess)
                {
                    lblResult.Text = "The email address is valid.";
                }
                else
                {
                    lblResult.Text = "The mail address is invalid,return code is : " + result.ReturnCode + ".";
                }
            }
            catch (Exception ex)
            {
                lblResult.Text += "<br>" + ex.Message;
            }
            // ExEnd:EmailValidations
        }
    }
}
