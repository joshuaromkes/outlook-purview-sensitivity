using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Outlook_Purview_Sensitivity
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                if (this.Application == null) return;

                Outlook.Explorer explorer = this.Application.ActiveExplorer();
                if (explorer != null)
                {
                    WireUpExplorer(explorer);
                }
                else if (this.Application.Explorers != null)
                {
                    this.Application.Explorers.NewExplorer += Explorers_NewExplorer;
                }
            }
            catch (Exception)
            {
                // Silently continue — column will be added on first folder switch
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Explorers_NewExplorer(Outlook.Explorer explorer)
        {
            WireUpExplorer(explorer);
        }

        private void WireUpExplorer(Outlook.Explorer explorer)
        {
            if (explorer == null) return;

            explorer.FolderSwitch += Explorer_FolderSwitch;

            Outlook.MAPIFolder folder = explorer.CurrentFolder;
            if (folder != null)
            {
                ColumnManager.EnsureColumn(folder);
                Marshal.ReleaseComObject(folder);
            }
        }

        private void Explorer_FolderSwitch()
        {
            Outlook.Explorer explorer = this.Application?.ActiveExplorer();
            if (explorer == null) return;

            Outlook.MAPIFolder folder = explorer.CurrentFolder;
            if (folder == null) return;

            ColumnManager.EnsureColumn(folder);
            ColumnManager.StampFolder(folder, maxItems: 50);
            Marshal.ReleaseComObject(folder);
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
