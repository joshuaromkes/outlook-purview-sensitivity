using System.Text.RegularExpressions;

namespace OutlookPurviewColumn
{
    /// <summary>
    /// Parses the raw msip_labels MAPI property string to extract
    /// the human-readable sensitivity label name.
    ///
    /// The msip_labels string format:
    ///   MSIP_Label_{GUID}_Enabled=true;
    ///   MSIP_Label_{GUID}_SetDate=2026-06-25T14:00:00Z;
    ///   MSIP_Label_{GUID}_Method=Privileged;
    ///   MSIP_Label_{GUID}_Name=PII High;
    ///   MSIP_Label_{GUID}_SiteId={tenant-guid};
    ///   ... (additional labels if multiple applied)
    ///
    /// We extract the _Name= value from the first Enabled=true label.
    /// </summary>
    internal static class LabelResolver
    {
        /// <summary>
        /// Regex to capture the label name from the msip_labels string.
        /// Matches: MSIP_Label_{guid}_Name={name};
        /// Captures: the label name (everything before the semicolon after _Name=)
        /// </summary>
        private static readonly Regex LabelNameRegex = new Regex(
            @"MSIP_Label_[0-9a-fA-F\-]+_Name=([^;]+)",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        /// <summary>
        /// Extracts the sensitivity label display name from a raw msip_labels string.
        /// </summary>
        /// <param name="msipLabels">Raw msip_labels MAPI property value (can be null).</param>
        /// <returns>The label name (e.g., "PII High"), or "None" if no label found.</returns>
        public static string GetLabelName(string msipLabels)
        {
            if (string.IsNullOrWhiteSpace(msipLabels))
                return "None";

            var match = LabelNameRegex.Match(msipLabels);
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }

            return "None";
        }
    }
}
