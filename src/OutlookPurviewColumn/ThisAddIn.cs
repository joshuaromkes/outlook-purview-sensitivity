using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookPurviewColumn
{
    /// <summary>
    /// VSTO Add-in entry point.
    ///
    /// On startup: hooks Explorer.FolderSwitch to inject the PurviewLabel column
    /// into every folder the user navigates to, and stamps visible items.
    ///
    /// On new mail: hooks Items.ItemAdd on the inbox and sent items folders
    /// to stamp incoming/outgoing messages in real time.
    /// </summary>
    public partial class ThisAddIn
    {
        /// <summary>
        /// Track which folders we've already hooked ItemAdd on,
        /// so we don't double-subscribe.
        /// </summary>
        private readonly Dictionary<string, Items> _hookedFolders = new Dictionary<string, Items>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Debug.WriteLine("[OutlookPurviewColumn] Add-in starting up...");

            // Hook folder switches — inject column when user navigates
            Application.ActiveExplorer().FolderSwitch += Explorer_FolderSwitch;

            // Hook the current folder immediately (Outlook may already have one open)
            var currentFolder = Application.ActiveExplorer().CurrentFolder;
            if (currentFolder != null)
            {
                OnFolderOpened(currentFolder);
                Marshal.ReleaseComObject(currentFolder);
            }

            Debug.WriteLine("[OutlookPurviewColumn] Add-in started successfully.");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Debug.WriteLine("[OutlookPurviewColumn] Add-in shutting down...");

            // Unhook all folder item listeners
            foreach (var kvp in _hookedFolders)
            {
                kvp.Value.ItemAdd -= Items_ItemAdd;
                Marshal.ReleaseComObject(kvp.Value);
            }
            _hookedFolders.Clear();

            Application.ActiveExplorer().FolderSwitch -= Explorer_FolderSwitch;
        }

        /// <summary>
        /// Fires when the user switches to a different folder in Outlook.
        /// Ensures the column exists and stamps items.
        /// </summary>
        private void Explorer_FolderSwitch()
        {
            try
            {
                var folder = Application.ActiveExplorer().CurrentFolder;
                if (folder != null)
                {
                    OnFolderOpened(folder);
                    Marshal.ReleaseComObject(folder);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[OutlookPurviewColumn] Error on folder switch: {ex.Message}");
            }
        }

        /// <summary>
        /// Called when a folder is opened (initial load or folder switch).
        /// Injects the PurviewLabel column and batch-stamps items.
        /// </summary>
        private void OnFolderOpened(MAPIFolder folder)
        {
            try
            {
                // Step 1: Ensure the PurviewLabel column exists in this folder's view
                ColumnManager.EnsureColumn(folder);

                // Step 2: Hook ItemAdd for real-time stamping (if not already hooked)
                HookItemAdd(folder);

                // Step 3: Batch-stamp existing items so they show labels immediately
                ColumnManager.StampFolder(folder, maxItems: 100);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[OutlookPurviewColumn] Error opening folder '{folder.Name}': {ex.Message}");
            }
        }

        /// <summary>
        /// Subscribes to Items.ItemAdd for the given folder so new messages
        /// get stamped in real time.
        /// </summary>
        private void HookItemAdd(MAPIFolder folder)
        {
            try
            {
                string key = folder.EntryID;
                if (_hookedFolders.ContainsKey(key))
                    return; // Already hooked

                var items = folder.Items;
                items.ItemAdd += Items_ItemAdd;
                _hookedFolders[key] = items;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[OutlookPurviewColumn] Error hooking ItemAdd for folder '{folder.Name}': {ex.Message}");
            }
        }

        /// <summary>
        /// Fires when a new item arrives in a hooked folder.
        /// Stamps it with its sensitivity label immediately.
        /// </summary>
        private void Items_ItemAdd(object item)
        {
            try
            {
                if (item is MailItem mailItem)
                {
                    ColumnManager.StampItem(mailItem);
                    Marshal.ReleaseComObject(mailItem);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[OutlookPurviewColumn] Error on ItemAdd: {ex.Message}");
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
