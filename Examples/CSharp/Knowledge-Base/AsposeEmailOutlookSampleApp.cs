using Aspose.Email.Mapi;
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

namespace Aspose.Email.Examples.CSharp.Email.Knowledge.Base
{
    public partial class AsposeEmailOutlookSampleApp : Form
    {
        public AsposeEmailOutlookSampleApp()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // ExStart:AsposeEmailOutlook

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_KnowledgeBase();

            // Load the outlook message file
            MapiMessage msg1 = MapiMessage.FromFile(dataDir + "Message.msg");

            // Use the public properties to assign the values to controls
            txtSubject.Text = msg1.Subject;
            txtFrom.Text = msg1.SenderEmailAddress;
            txtTo.Text = msg1.DisplayTo;
            txtBody.Text = msg1.Body;
            // ExEnd:AsposeEmailOutlook
        }
    }
}
