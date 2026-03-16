using System;
using System.Diagnostics;

namespace SaveAsPDF.Helpers
{
    /// <summary>
    /// Determines whether the main task pane should react to Outlook events.
    /// </summary>
    internal static class MainPaneBehavior
    {
        public static bool ShouldRefreshOnSelectionChange(bool hasMainControl, bool hasMainPane, bool isMainPaneVisible)
        {
            return hasMainControl && hasMainPane && isMainPaneVisible;
        }

        public static bool ToggleVisibility(bool currentVisibility)
        {
            return !currentVisibility;
        }

        public static bool ShouldLoadMailItemAfterToggle(bool newVisibility)
        {
            return newVisibility;
        }
    }

    /// <summary>
    /// Lightweight diagnostics logger used throughout the add-in.
    /// </summary>
    internal static class DiagnosticsLogger
    {
        public static void LogException(string context, Exception ex)
        {
            var message = string.Format(
                "[{0}] {1}: {2}",
                context ?? "Unknown",
                ex == null ? "Exception" : ex.GetType().FullName,
                ex == null ? "null" : ex.Message);
            Trace.WriteLine(message);
            Debug.WriteLine(message);
        }
    }
}
