using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using SysException = System.Exception;

namespace Outlook_Purview_Sensitivity
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
                UserDefinedProperties props = folder.UserDefinedProperties;
                bool found = false;
                for (int i = 1; i <= props.Count; i++)
                {
                    UserDefinedProperty prop = props[i];
                    if (prop.Name == FieldName)
                    {
                        found = true;
                        Marshal.ReleaseComObject(prop);
                        break;
                    }
                    Marshal.ReleaseComObject(prop);
                }

                if (!found)
                    props.Add(FieldName, OlUserPropertyType.olText);

                Marshal.ReleaseComObject(props);

                // Step 2: Add field to the table view (if missing)
                object view = folder.CurrentView;
                if (view is TableView tableView)
                {
                    ViewFields fields = tableView.ViewFields;
                    bool columnFound = false;
                    for (int i = 1; i <= fields.Count; i++)
                    {
                        ViewField field = fields[i];
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
            catch (SysException)
            {
                // Non-critical — column will be added on next folder switch
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
                // HARDCODED TEST: verifies stamping pipeline works
                string labelName = "TEST-LABEL";

                // Check if already stamped with the correct value
                try
                {
                    UserProperty existing = mailItem.UserProperties[FieldName];
                    if (existing != null)
                    {
                        string val = existing.Value as string;
                        Marshal.ReleaseComObject(existing);
                        if (val == labelName) return; // Already correct
                    }
                }
                catch
                {
                    // Property doesn't exist yet — expected
                }

                // Add or update
                UserProperty userProp = mailItem.UserProperties.Add(
                    FieldName, OlUserPropertyType.olText);
                userProp.Value = labelName;
                mailItem.Save();
                Marshal.ReleaseComObject(userProp);
            }
            catch (SysException)
            {
                // Skip items that can't be stamped
            }
        }

        /// <summary>
        /// Batch stamps visible items in the folder view (up to maxItems).
        /// </summary>
        public static void StampFolder(MAPIFolder folder, int maxItems = 50)
        {
            if (folder == null) return;

            try
            {
                Items items = folder.Items;
                int count = 0;
                for (int i = 1; i <= items.Count && count < maxItems; i++)
                {
                    object item = items[i];
                    if (item is MailItem mailItem)
                    {
                        StampItem(mailItem);
                        count++;
                        Marshal.ReleaseComObject(mailItem);
                    }
                    Marshal.ReleaseComObject(item);
                }
                Marshal.ReleaseComObject(items);
            }
            catch (SysException)
            {
                // Non-critical
            }
        }
    }
}
