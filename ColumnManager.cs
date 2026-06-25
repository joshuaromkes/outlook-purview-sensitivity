using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using SysException = System.Exception;

namespace Outlook_Purview_Sensitivity
{
    internal static class ColumnManager
    {
        public const string FieldName = "PurviewLabel";

        public static void EnsureColumn(MAPIFolder folder)
        {
            if (folder == null) return;

            try
            {
                Debug.WriteLine($"[CM] EnsureColumn: {folder.Name}");

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
                {
                    props.Add(FieldName, OlUserPropertyType.olText);
                    Debug.WriteLine("[CM] Added user-defined property");
                }

                Marshal.ReleaseComObject(props);

                object view = folder.CurrentView;
                Debug.WriteLine($"[CM] View type: {view?.GetType().Name ?? "null"}");

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
                        Debug.WriteLine("[CM] Added column to view");
                    }
                    else
                    {
                        Debug.WriteLine("[CM] Column already in view");
                    }

                    Marshal.ReleaseComObject(fields);
                    Marshal.ReleaseComObject(tableView);
                }

                Marshal.ReleaseComObject(view);
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[CM] EnsureColumn error: {ex.Message}");
            }
        }

        public static void StampItem(MailItem mailItem)
        {
            if (mailItem == null) return;

            try
            {
                string labelName = "TEST-LABEL";

                try
                {
                    UserProperty existing = mailItem.UserProperties[FieldName];
                    if (existing != null)
                    {
                        string val = existing.Value as string;
                        Marshal.ReleaseComObject(existing);
                        if (val == labelName) return;
                    }
                }
                catch
                {
                }

                UserProperty userProp = mailItem.UserProperties.Add(
                    FieldName, OlUserPropertyType.olText);
                userProp.Value = labelName;
                mailItem.Save();
                Marshal.ReleaseComObject(userProp);

                Debug.WriteLine($"[CM] Stamped: {mailItem.Subject}");
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[CM] StampItem error: {ex.Message}  Subject: {mailItem.Subject}");
            }
        }

        public static void StampFolder(MAPIFolder folder, int maxItems = 50)
        {
            if (folder == null) return;

            try
            {
                Items items = folder.Items;
                Debug.WriteLine($"[CM] StampFolder: {folder.Name}  Items.Count={items.Count}");

                int count = 0;
                int mailCount = 0;
                for (int i = 1; i <= items.Count && count < maxItems; i++)
                {
                    object item = items[i];
                    if (item is MailItem mailItem)
                    {
                        mailCount++;
                        StampItem(mailItem);
                        count++;
                        Marshal.ReleaseComObject(mailItem);
                    }
                    Marshal.ReleaseComObject(item);
                }

                Debug.WriteLine($"[CM] StampFolder done: {count} stamped, {mailCount} mail items seen");
                Marshal.ReleaseComObject(items);

                // Force view refresh to show new values
                object view = folder.CurrentView;
                if (view is TableView tableView)
                {
                    tableView.Apply();
                    Debug.WriteLine("[CM] View refreshed");
                    Marshal.ReleaseComObject(tableView);
                }
                Marshal.ReleaseComObject(view);
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[CM] StampFolder error: {ex.Message}");
            }
        }
    }
}
