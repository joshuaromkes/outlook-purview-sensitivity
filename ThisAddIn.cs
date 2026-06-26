using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using SysException = System.Exception;

namespace Outlook_Purview_Sensitivity
{
    public partial class ThisAddIn
    {
        private Outlook.Items _currentItems;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Debug.WriteLine("[PS] ========================================");
                Debug.WriteLine("[PS] Outlook-Purview-Sensitivity starting up");
                Debug.WriteLine($"[PS] Add-in version: {GetType().Assembly.GetName().Version}");

                if (this.Application == null)
                {
                    Debug.WriteLine("[PS] HEALTH: FAIL — Application object is null, add-in cannot function");
                    return;
                }

                Debug.WriteLine("[PS] HEALTH: OK — Outlook Application object found");
                Debug.WriteLine($"[PS] Outlook version: {this.Application.Version}");

                Outlook.Explorer explorer = null;
                try
                {
                    explorer = this.Application.ActiveExplorer();
                    if (explorer != null)
                    {
                        WireUpExplorer(explorer);
                    }
                    else
                    {
                        Debug.WriteLine("[PS] No active explorer window — waiting for NewExplorer event");
                    }
                }
                catch (SysException ex)
                {
                    Debug.WriteLine($"[PS] ActiveExplorer error: {ex.Message}");
                }
                finally
                {
                    if (explorer != null) Marshal.ReleaseComObject(explorer);
                }

                this.Application.Explorers.NewExplorer += Explorers_NewExplorer;

                Debug.WriteLine("[PS] Startup complete — listening for explorer events");
                Debug.WriteLine("[PS] ========================================");
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[PS] HEALTH: FAIL — Startup error: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Debug.WriteLine("[PS] Add-in shutting down");
        }

        private void Explorers_NewExplorer(Outlook.Explorer explorer)
        {
            if (explorer == null) return;
            WireUpExplorer(explorer);
        }

        private void WireUpExplorer(Outlook.Explorer explorer)
        {
            if (explorer == null) return;

            try
            {
                explorer.FolderSwitch += Explorer_FolderSwitch;

                Outlook.MAPIFolder folder = null;
                try
                {
                    folder = explorer.CurrentFolder;
                    if (folder != null)
                    {
                        Debug.WriteLine($"[PS] WireUp: {folder.Name}");

                        ColumnManager.EnsureColumn(folder);

                        DetachItemsHandler();
                        _currentItems = folder.Items;
                        _currentItems.ItemAdd += Items_ItemAdd;

                        ColumnManager.StampFolder(folder, maxItems: 50);
                    }
                }
                catch (SysException ex)
                {
                    Debug.WriteLine($"[PS] WireUp folder processing error: {ex.Message}");
                }
                finally
                {
                    if (folder != null) Marshal.ReleaseComObject(folder);
                }
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[PS] WireUpExplorer error: {ex.Message}");
            }
        }

        private void Explorer_FolderSwitch()
        {
            try
            {
                Outlook.Explorer explorer = null;
                Outlook.MAPIFolder folder = null;

                try
                {
                    explorer = this.Application?.ActiveExplorer();
                    if (explorer == null) return;

                    folder = explorer.CurrentFolder;
                    if (folder == null) return;

                    Debug.WriteLine($"[PS] FolderSwitch: {folder.Name}");

                    ColumnManager.EnsureColumn(folder);

                    DetachItemsHandler();
                    _currentItems = folder.Items;
                    _currentItems.ItemAdd += Items_ItemAdd;

                    ColumnManager.StampFolder(folder, maxItems: 50);
                }
                finally
                {
                    if (folder != null) Marshal.ReleaseComObject(folder);
                    if (explorer != null) Marshal.ReleaseComObject(explorer);
                }
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[PS] FolderSwitch error: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private void DetachItemsHandler()
        {
            if (_currentItems != null)
            {
                try
                {
                    _currentItems.ItemAdd -= Items_ItemAdd;
                }
                catch (SysException ex)
                {
                    Debug.WriteLine($"[PS] DetachItemsHandler error: {ex.Message}");
                }
                finally
                {
                    Marshal.ReleaseComObject(_currentItems);
                    _currentItems = null;
                }
            }
        }

        private void Items_ItemAdd(object item)
        {
            try
            {
                if (item is Outlook.MailItem mailItem)
                {
                    string label = LabelReader.GetLabelName(mailItem);
                    if (label != "None")
                    {
                        Debug.WriteLine($"[PS] ItemAdd: new labeled item, stamping");
                        ColumnManager.StampItem(mailItem);
                    }
                    Marshal.ReleaseComObject(mailItem);
                }
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[PS] Items_ItemAdd error: {ex.Message}");
            }
            finally
            {
                if (item != null) Marshal.ReleaseComObject(item);
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
