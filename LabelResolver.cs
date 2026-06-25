namespace Outlook_Purview_Sensitivity
{
    /// <summary>
    /// Parses the raw msip_labels MAPI property string to extract
    /// the human-readable sensitivity label name.
    ///
    /// The msip_labels string format:
    ///   MSIP_Label_{guid}_Enabled=true;
    ///   MSIP_Label_{guid}_SetDate=2026-06-25T14:00:00Z;
    ///   MSIP_Label_{guid}_Method=Privileged;
    ///   MSIP_Label_{guid}_Name=PII High;
    ///   MSIP_Label_{guid}_SiteId={tenant-guid};
    ///
    /// We extract the _Name= value from the string.
    /// </summary>
    internal static class LabelResolver
    {
        /// <summary>
        /// Extracts the sensitivity label display name from a raw msip_labels value.
        /// </summary>
        /// <param name="msipLabels">Raw msip_labels MAPI property value (can be null).</param>
        /// <returns>The label name (e.g., "PII High"), or "None" if not found.</returns>
        public static string GetLabelName(string msipLabels)
        {
            if (string.IsNullOrWhiteSpace(msipLabels))
                return "None";

            // Find the first _Name= occurrence
            const string nameMarker = "_Name=";
            int nameIdx = msipLabels.IndexOf(nameMarker, System.StringComparison.Ordinal);
            if (nameIdx < 0)
                return "None";

            // Skip past "_Name="
            int valueStart = nameIdx + nameMarker.Length;

            // Find the semicolon that terminates the value
            int valueEnd = msipLabels.IndexOf(';', valueStart);
            if (valueEnd < 0)
                valueEnd = msipLabels.Length;

            string labelName = msipLabels.Substring(valueStart, valueEnd - valueStart).Trim();
            return string.IsNullOrEmpty(labelName) ? "None" : labelName;
        }
    }
}
