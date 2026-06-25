using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using SysException = System.Exception;

namespace OutlookPurviewColumn
{
    /// <summary>
    /// Manages the custom "PurviewLabel" column in Outlook folder views.
    /// All COM interop collections are iterated with integer-indexed loops
    /// to avoid foreach/enumerator issues with Office interop.
    /// </summary>
    internal static class ColumnManager
    {
        public const string FieldName = "PurviewLabel";

        /// <summary>
        /// Ensures the PurviewLabel user-defined property exists in the folder
        /// and is added as a column in the current table view.
        /// </summary>
        public static void EnsureColumn(MAPIFolder folder)
        {
            if (folder == null) return;

            try
            {
                // Step 1: Add user-defined property to folder (if missing)
                var props = folder.UserDefinedProperties;
                bool found = false;
                for (int i = 1; i <= props.Count; i++)
                {
                    var prop = props[i];
                    if (prop.Name == FieldName)
                    {
                        found = true;
                        Marshal.ReleaseComObject(prop);
                        break;
                    }
                    Marshal.ReleaseComObject(prop);
                }

                if (!found)
                {
                    props.Add(FieldName, OlUserPropertyType.olText);
                }
                Marshal.ReleaseComObject(props);

                // Step 2: Add field to table view (if missing)
                var view = folder.CurrentView;
                if (view is TableView tableView)
                {
                    var fields = tableView.ViewFields;
                    bool columnFound = false;
                    for (int i = 1; i <= fields.Count; i++)
                    {
                        var field = fields[i];
                        if (field.ViewXMLSchemaName == FieldName)
                        {
                            columnFound = true;
                            Marshal.ReleaseComObject(field);
                            break;
                        }
                        Marshal.ReleaseComObject(field);
                    }

                    if (!columnFound)
                    {
                        fields.Add(FieldName);
                        tableView.Save();
                    }
                    Marshal.ReleaseComObject(fields);
                    Marshal.ReleaseComObject(tableView);
                }
                Marshal.ReleaseComObject(view);
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[OutlookPurviewColumn] EnsureColumn error: {ex.Message}");
            }
        }

        /// <summary>
        /// Stamps a single MailItem with its Purview label.
        /// </summary>
        public static void StampItem(MailItem mailItem)
        {
            if (mailItem == null) return;

            try
            {
                var labelName = LabelReader.GetLabelName(mailItem);

                // Check if already stamped with correct value
                try
                {
                    var existing = mailItem.UserProperties[FieldName];
                    if (existing != null)
                    {
                        var val = existing.Value as string;
                        Marshal.ReleaseComObject(existing);
                        if (val == labelName) return; // Already correct
                    }
                }
                catch
                {
                    // Property doesn't exist yet — expected
                }

                // Add or update
                var userProp = mailItem.UserProperties.Add(FieldName, OlUserPropertyType.olText);
                userProp.Value = labelName;
                mailItem.Save();
                Marshal.ReleaseComObject(userProp);
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[OutlookPurviewColumn] StampItem error: {ex.Message}");
            }
        }

        /// <summary>
        /// Batch stamps visible items in the folder view.
        /// </summary>
        public static void StampFolder(MAPIFolder folder, int maxItems = 50)
        {
            if (folder == null) return;

            try
            {
                var items = folder.Items;
                int count = 0;
                for (int i = 1; i <= items.Count && count < maxItems; i++)
                {
                    var item = items[i];
                    if (item is MailItem mailItem)
                    {
                        // Skip items without labels to avoid stamping "None" on everything
                        var label = LabelReader.GetLabelName(mailItem);
                        if (label != "None")
                        {
                            StampItem(mailItem);
                            count++;
                        }
                        Marshal.ReleaseComObject(mailItem);
                    }
                    Marshal.ReleaseComObject(item);
                }
                Marshal.ReleaseComObject(items);
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[OutlookPurviewColumn] StampFolder error: {ex.Message}");
            }
        }
    }
}
