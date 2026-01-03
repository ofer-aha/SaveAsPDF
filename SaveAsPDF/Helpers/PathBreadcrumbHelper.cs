using System;

namespace SaveAsPDF.Helpers
{
    public static class PathBreadcrumbHelper
    {
        /// <summary>
        /// Formats a filesystem path into a breadcrumb-style string using the explorer-like separator (›).
        /// </summary>
        public static string FormatPathAsBreadcrumb(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return string.Empty;

            // Remove trailing backslash and split
            var trimmed = path.TrimEnd('\\');
            var segments = trimmed.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" › ", segments);
        }

        /// <summary>
        /// Converts a breadcrumb-style string back into a filesystem path.
        /// </summary>
        public static string ToPathFromBreadcrumb(string breadcrumb)
        {
            if (string.IsNullOrWhiteSpace(breadcrumb))
                return string.Empty;

            return breadcrumb.Contains(" › ") ? breadcrumb.Replace(" › ", "\\") : breadcrumb;
        }
    }
}
