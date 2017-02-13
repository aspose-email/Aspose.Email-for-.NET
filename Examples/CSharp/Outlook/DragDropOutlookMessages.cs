using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    public partial class DragDropOutlookMessages : Form
    {
        public DragDropOutlookMessages()
        {
            InitializeComponent();
        }

        // ExStart:DisplayFilesCount
        private void myPanel1_DragDrop(object sender, DragEventArgs e)
        {

            Aspose.Email.Windows.Forms.FileDragEventArgs args;
            args = (Aspose.Email.Windows.Forms.FileDragEventArgs)e;
            if (args != null && args.Files.Count > 0)
            {
                for (int i = 0; i < args.Files.Count; i++)
                {
                    SaveFileDialog dialog = new SaveFileDialog();
                    dialog.FileName = args.Files[i].FileName;
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            System.IO.FileStream output;
                            output = new System.IO.FileStream(dialog.FileName, System.IO.FileMode.CreateNew);
                            args.Files[i].Save(output);
                            MessageBox.Show("Save success:" + dialog.FileName);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Save failed:" + ex.ToString());
                        }
                    }
                }
            }
        }
        // ExEnd:DisplayFilesCount
    }
}
