using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    // ExStart:FileDropTargetManager
    class MyPanel : System.Windows.Forms.Panel	
    {
        private Aspose.Email.Windows.Forms.FileDropTargetManager manager;

        // Hook up
        protected override void OnHandleCreated(EventArgs e)
        {
            this.manager = new Aspose.Email.Windows.Forms.FileDropTargetManager(this);
            this.manager.EnsureRegistered(this);
            base.OnHandleCreated(e);
        }

        // Unhook
        protected override void OnHandleDestroyed(EventArgs e)
        {
            this.manager.EnsureUnRegistered(this);
            base.OnHandleDestroyed(e);
        }
    }
    // ExEnd:FileDropTargetManager
}
