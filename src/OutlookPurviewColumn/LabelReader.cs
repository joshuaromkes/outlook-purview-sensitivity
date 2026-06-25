using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Outlook;

namespace OutlookPurviewColumn
{
    /// <summary>
    /// Reads the Purview sensitivity label from a MailItem via MAPI PropertyAccessor.
    ///
    /// The msip_labels x-header is stored as a named MAPI string property under
    /// the PS_INTERNET_HEADERS property set ({00020386-0000-0000-C000-000000000046}).
    ///
    /// Property path:
    ///   http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/msip_labels/0x0000001F
    /// </summary>
    internal static class LabelReader
    {
        /// <summary>
        /// Full DASL property path for the msip_labels named MAPI property.
        /// </summary>
        private const string MsipLabelsPropertyPath =
            "http://schemas.microsoft.com/mapi/string/" +
            "{00020386-0000-0000-C000-000000000046}/msip_labels/0x0000001F";

        /// <summary>
        /// Reads the sensitivity label display name from a MailItem.
        /// </summary>
        /// <param name="mailItem">The Outlook MailItem to read from.</param>
        /// <returns>The label name (e.g., "PII High"), or "None" if no label or error.</returns>
        public static string GetLabelName(MailItem mailItem)
        {
            if (mailItem == null)
                return "None";

            try
            {
                // Attempt to read the msip_labels MAPI property
                var msipLabels = mailItem.PropertyAccessor.GetProperty(MsipLabelsPropertyPath) as string;

                // Parse the friendly name from the msip string
                return LabelResolver.GetLabelName(msipLabels);
            }
            catch (Exception ex)
            {
                // Property may not exist on the item (no label applied, or not an email
                // with Purview labels). This is an expected case — not an error.
                Debug.WriteLine($"[OutlookPurviewColumn] Could not read label: {ex.Message}");
                return "None";
            }
        }
    }
}
