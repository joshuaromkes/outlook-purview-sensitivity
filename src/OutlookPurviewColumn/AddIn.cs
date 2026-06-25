using System;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Interop.Outlook;

namespace OutlookPurviewColumn
{
    [ComVisible(true)]
    [Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]
    [ProgId("OutlookPurviewColumn.AddIn")]
    public class AddIn : IDTExtensibility2
    {
        private object _application;

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _application = Application;
            var outlookApp = Application as Microsoft.Office.Interop.Outlook.Application;
            if (outlookApp == null) return;

            // Hook folder switches
            var explorer = outlookApp.ActiveExplorer();
            explorer.FolderSwitch += Explorer_FolderSwitch;

            // Handle current folder if Outlook already has one open
            var currentFolder = explorer.CurrentFolder;
            if (currentFolder != null)
            {
                OnFolderOpened(currentFolder);
                Marshal.ReleaseComObject(currentFolder);
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            var outlookApp = _application as Microsoft.Office.Interop.Outlook.Application;
            if (outlookApp != null)
            {
                outlookApp.ActiveExplorer().FolderSwitch -= Explorer_FolderSwitch;
            }
            _application = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        private void Explorer_FolderSwitch()
        {
            try
            {
                var outlookApp = _application as Microsoft.Office.Interop.Outlook.Application;
                if (outlookApp == null) return;

                var folder = outlookApp.ActiveExplorer().CurrentFolder;
                if (folder != null)
                {
                    OnFolderOpened(folder);
                    Marshal.ReleaseComObject(folder);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OutlookPurviewColumn] Error on folder switch: {ex.Message}");
            }
        }

        private void OnFolderOpened(MAPIFolder folder)
        {
            try
            {
                ColumnManager.EnsureColumn(folder);
                ColumnManager.StampFolder(folder, maxItems: 100);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OutlookPurviewColumn] Error opening folder: {ex.Message}");
            }
        }
    }
}
