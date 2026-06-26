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
                    Debug.WriteLine($"[CM] {folder.Name}: added user property '{FieldName}'");
                }

                Marshal.ReleaseComObject(props);

                EnsureColumnInView(folder);
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[CM] {folder.Name}: EnsureColumn error: {ex.Message}");
            }
        }

        private static void EnsureColumnInView(MAPIFolder folder)
        {
            object view = null;
            TableView tableView = null;

            try
            {
                view = folder.CurrentView;

                tableView = view as TableView;
                if (tableView == null)
                {
                    Debug.WriteLine($"[CM] {folder.Name}: current view '{folder.CurrentView.Name}' is not TableView — switching to Messages");
                    Marshal.ReleaseComObject(view);
                    view = null;
                    SwitchToMessagesView(folder);
                    view = folder.CurrentView;
                    tableView = view as TableView;
                }

                if (tableView != null)
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
                        ViewField newField = fields.Add(FieldName);
                        Marshal.ReleaseComObject(newField);
                        tableView.Save();
                        Debug.WriteLine($"[CM] {folder.Name}: added column '{FieldName}' to view");
                    }

                    Marshal.ReleaseComObject(fields);
                    Marshal.ReleaseComObject(tableView);
                    tableView = null;
                    view = null;
                }
                else
                {
                    Debug.WriteLine($"[CM] {folder.Name}: cannot get TableView even after switch — view type is {view?.GetType().Name}");
                }
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[CM] {folder.Name}: EnsureColumnInView error: {ex.Message}");
            }
            finally
            {
                if (tableView != null) Marshal.ReleaseComObject(tableView);
                if (view != null) Marshal.ReleaseComObject(view);
            }
        }

        public static void StampItem(MailItem mailItem)
        {
            if (mailItem == null) return;

            try
            {
                string labelName = LabelReader.GetLabelName(mailItem);

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
                Marshal.ReleaseComObject(userProp);

                SaveMailItem(mailItem);
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[CM] StampItem error: {ex.Message}");
            }
        }

        private static void SaveMailItem(MailItem mailItem)
        {
            try
            {
                mailItem.Save();
            }
            catch (SysException ex)
            {
                Debug.WriteLine(
                    $"[CM] Save failed (shared/delegate mailbox, PST, or read-only): {ex.GetType().Name}: {ex.Message}");
            }
        }

        public static void StampFolder(MAPIFolder folder, int maxItems = 50)
        {
            if (folder == null) return;

            try
            {
                Items items = folder.Items;
                int stampCount = 0;
                int mailCount = 0;
                for (int i = 1; i <= items.Count && stampCount < maxItems; i++)
                {
                    object item = null;
                    try
                    {
                        item = items[i];
                        if (item is MailItem mailItem)
                        {
                            mailCount++;
                            string label = LabelReader.GetLabelName(mailItem);
                            if (label != "None")
                            {
                                StampItem(mailItem);
                                stampCount++;
                            }
                            Marshal.ReleaseComObject(mailItem);
                        }
                    }
                    catch (SysException ex)
                    {
                        Debug.WriteLine($"[CM] {folder.Name}: skip item[{i}]: {ex.Message}");
                    }
                    finally
                    {
                        if (item != null) Marshal.ReleaseComObject(item);
                    }
                }
                Marshal.ReleaseComObject(items);

                Debug.WriteLine(
                    $"[CM] {folder.Name}: stamped {stampCount}/{mailCount} labeled items");

                RefreshView(folder);
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[CM] {folder.Name}: StampFolder error: {ex.Message}");
            }
        }

        private static void RefreshView(MAPIFolder folder)
        {
            object view = null;
            TableView tableView = null;

            try
            {
                view = folder.CurrentView;
                tableView = view as TableView;

                if (tableView == null)
                {
                    Debug.WriteLine(
                        $"[CM] {folder.Name}: view '{folder.CurrentView.Name}' not TableView — switch to Messages for refresh");
                    Marshal.ReleaseComObject(view);
                    view = null;
                    SwitchToMessagesView(folder);
                    view = folder.CurrentView;
                    tableView = view as TableView;
                }

                if (tableView != null)
                {
                    tableView.Apply();
                    Marshal.ReleaseComObject(tableView);
                    tableView = null;
                    view = null;
                }
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[CM] {folder.Name}: RefreshView error: {ex.Message}");
            }
            finally
            {
                if (tableView != null) Marshal.ReleaseComObject(tableView);
                if (view != null) Marshal.ReleaseComObject(view);
            }
        }

        private static void SwitchToMessagesView(MAPIFolder folder)
        {
            Views views = null;
            object messagesView = null;

            try
            {
                views = folder.Views;

                for (int i = 1; i <= views.Count; i++)
                {
                    View v = views[i];
                    if (v.Name == "Messages")
                    {
                        messagesView = v;
                        Marshal.ReleaseComObject(v);
                        break;
                    }
                    Marshal.ReleaseComObject(v);
                }

                if (messagesView != null)
                {
                    ((View)messagesView).Apply();
                }
                else
                {
                    View msgView = views["Messages"];
                    if (msgView != null)
                        msgView.Apply();
                    else
                        Debug.WriteLine($"[CM] {folder.Name}: Messages view not found");
                }

                Debug.WriteLine($"[CM] {folder.Name}: set CurrentView to Messages");
            }
            catch (SysException ex)
            {
                Debug.WriteLine($"[CM] {folder.Name}: SwitchToMessagesView failed: {ex.Message}");
            }
            finally
            {
                if (messagesView != null) Marshal.ReleaseComObject(messagesView);
                if (views != null) Marshal.ReleaseComObject(views);
            }
        }
    }
}
