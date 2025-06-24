using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using SaveAsPDF.Helpers;

namespace SaveAsPDF
{
    public partial class ThisAddIn
    {
        public static Outlook.MailItem TypeOfMailitem(Outlook.MailItem mailItem)
        {
            mailItem = null;
            dynamic windowType = Globals.ThisAddIn.Application.ActiveWindow();
            if (windowType is Outlook.Explorer)
            {
                // frmMain Explorer
                Outlook.Explorer explorer = windowType as Outlook.Explorer;
                if (explorer.Selection.Count > 0)
                {
                    mailItem = explorer.Selection[1] as Outlook.MailItem;
                }
            }
            else if (windowType is Outlook.Inspector)
            {
                // Read or Compose
                Outlook.Inspector inspector = windowType as Outlook.Inspector;
                mailItem = inspector.CurrentItem as Outlook.MailItem;
            }

            return mailItem;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Initialize the add-in
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Clean up resources before shutdown
            CleanupResources();
        }

        /// <summary>
        /// Performs application cleanup to release resources
        /// </summary>
        private void CleanupResources()
        {
            try
            {
                // Release Office COM resources
                OfficeHelpers.ReleaseWordInstance();
                
                // Clear caches
                ComboBoxExtensions.ClearDirectoryCache();
                
                // Collect garbage to help release COM objects
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch
            {
                // Silently catch errors during cleanup
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}