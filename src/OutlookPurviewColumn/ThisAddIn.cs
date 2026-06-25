using System;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookPurviewColumn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var explorer = this.Application.ActiveExplorer();
            explorer.FolderSwitch += Explorer_FolderSwitch;

            var folder = explorer.CurrentFolder;
            if (folder != null)
            {
                ColumnManager.EnsureColumn(folder);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Explorer_FolderSwitch()
        {
            var folder = this.Application.ActiveExplorer().CurrentFolder;
            if (folder != null)
            {
                ColumnManager.EnsureColumn(folder);
                ColumnManager.StampFolder(folder, maxItems: 50);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
            }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
