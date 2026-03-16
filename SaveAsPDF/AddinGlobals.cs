using Outlook = Microsoft.Office.Interop.Outlook;

namespace SaveAsPDF
{
    /// <summary>
    /// Provides global access to the COM Add-in instance and the Outlook Application object.
    /// Replaces the VSTO-generated <c>Globals</c> class.
    /// </summary>
    internal static class AddinGlobals
    {
        /// <summary>
        /// The singleton COM Add-in entry point.
        /// </summary>
        internal static Connect Connect { get; set; }

        /// <summary>
        /// Shortcut to the host Outlook Application object.
        /// </summary>
        internal static Outlook.Application OutlookApp => Connect?.Application;
    }
}




