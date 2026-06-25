using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using SysException = System.Exception;

namespace OutlookPurviewColumn
{
    /// <summary>
    /// Manages the custom "PurviewLabel" column in Outlook folder views.
    ///
    /// Adds a user-defined field to each folder and injects it as a visible
    /// column in the folder's table view. Stamps individual MailItems with
    /// their sensitivity label so it appears in the column.
    /// </summary>
    internal static class ColumnManager
    {
        /// <summary>
        /// The name of the user-defined field we add to each folder.
        /// This is what appears as the column header.
        /// </summary>
        public const string FieldName = "PurviewLabel";

        /// <summary>
        /// Display name for the column in the view.
        /// </summary>
        public const string ColumnDisplayName = "Purview Label";

        /// <summary>
        /// Ensures the PurviewLabel column exists in the given folder's view.
        /// Idempotent — safe to call on every folder switch.
        /// </summary>
        /// <param name="folder">The Outlook folder to check/modify.</param>
        public static void EnsureColumn(MAPIFolder folder)
        {
            if (folder == null) return;

            try
            {
                // Step 1: Add the user-defined property to the folder (if missing)
                var props = folder.UserDefinedProperties;
                bool found = false;
                foreach (UserDefinedProperty prop in props)
                {
                    if (prop.Name == FieldName)
                    {
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    props.Add(FieldName, OlUserPropertyType.olText);
                    Debug.WriteLine($"[OutlookPurviewColumn] Added '{FieldName}' user-defined property to folder: {folder.Name}");
                }

                // Step 2: Add the field to the current view (if missing)
                var view = folder.CurrentView;
                if (view is TableView tableView)
                {
                    bool columnFound = false;
                    foreach (ViewField field in tableView.ViewFields)
                    {
                        if (field.ViewXMLSchemaName == FieldName)
                        {
                            columnFound = true;
                            break;
                        }
                    }

                    if (!columnFound)
                    {
                        tableView.ViewFields.Add(FieldName);
                        tableView.Save();
                        Debug.WriteLine($"[OutlookPurviewColumn] Added '{FieldName}' column to view: {folder.Name}");
                    }
                }
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[OutlookPurviewColumn] Error ensuring column in folder '{folder.Name}': {ex.Message}");
            }
        }

        /// <summary>
        /// Stamps a single MailItem with its Purview label in the UserProperties.
        /// This is what makes the label visible in the column.
        /// </summary>
        /// <param name="mailItem">The MailItem to stamp.</param>
        public static void StampItem(MailItem mailItem)
        {
            if (mailItem == null) return;

            try
            {
                var labelName = LabelReader.GetLabelName(mailItem);

                // Check if already stamped with the correct value
                try
                {
                    var existing = mailItem.UserProperties[FieldName];
                    if (existing != null && existing.Value as string == labelName)
                    {
                        Marshal.ReleaseComObject(existing);
                        return; // Already correct, skip
                    }
                    if (existing != null) Marshal.ReleaseComObject(existing);
                }
                catch
                {
                    // Property doesn't exist yet on this item — expected
                }

                // Stamp the label
                var userProp = mailItem.UserProperties.Add(FieldName, OlUserPropertyType.olText);
                userProp.Value = labelName;
                mailItem.Save();
                Marshal.ReleaseComObject(userProp);

                Debug.WriteLine($"[OutlookPurviewColumn] Stamped item: {mailItem.Subject} → {labelName}");
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[OutlookPurviewColumn] Error stamping item '{mailItem.Subject}': {ex.Message}");
            }
        }

        /// <summary>
        /// Batch stamps unlabeled items in a folder. Processes up to maxItems
        /// to avoid locking the UI on large folders.
        /// </summary>
        /// <param name="folder">The folder to scan.</param>
        /// <param name="maxItems">Maximum number of items to process in one batch.</param>
        public static void StampFolder(MAPIFolder folder, int maxItems = 100)
        {
            if (folder == null) return;

            try
            {
                var items = folder.Items;
                int count = 0;

                foreach (object item in items)
                {
                    if (count >= maxItems) break;

                    if (item is MailItem mailItem)
                    {
                        StampItem(mailItem);
                        Marshal.ReleaseComObject(mailItem);
                        count++;
                    }
                }

                Marshal.ReleaseComObject(items);

                if (count > 0)
                {
                    Debug.WriteLine($"[OutlookPurviewColumn] Batch stamped {count} items in folder: {folder.Name}");
                }
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[OutlookPurviewColumn] Error stamping folder '{folder.Name}': {ex.Message}");
            }
        }
    }
}
