using System;
using System.Reflection;
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
        private Microsoft.Office.Interop.Outlook.Application _app;

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _app = Application as Microsoft.Office.Interop.Outlook.Application;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            if (_app?.ActiveExplorer() != null)
            {
                _app.ActiveExplorer().FolderSwitch -= Explorer_FolderSwitch;
            }
            _app = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            // OnStartupComplete fires after the Outlook UI is fully loaded.
            // This is the safe time to access ActiveExplorer and hook events.
            try
            {
                if (_app == null) return;

                var explorer = _app.ActiveExplorer();
                if (explorer == null) return;

                explorer.FolderSwitch += Explorer_FolderSwitch;

                var folder = explorer.CurrentFolder;
                if (folder != null)
                {
                    ColumnManager.EnsureColumn(folder);
                    Marshal.ReleaseComObject(folder);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OutlookPurviewColumn] Startup error: {ex}");
            }
        }

        private void Explorer_FolderSwitch()
        {
            try
            {
                if (_app == null) return;
                var explorer = _app.ActiveExplorer();
                if (explorer == null) return;

                var folder = explorer.CurrentFolder;
                if (folder != null)
                {
                    ColumnManager.EnsureColumn(folder);
                    ColumnManager.StampFolder(folder, maxItems: 50);
                    Marshal.ReleaseComObject(folder);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OutlookPurviewColumn] FolderSwitch: {ex}");
            }
        }
    }
}
