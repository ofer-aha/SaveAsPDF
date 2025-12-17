using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Microsoft.Office.Interop.Outlook;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;

namespace SaveAsPDF
{
    public partial class ThisAddIn : ISettingsRequester
    {
        private CustomTaskPane _settingsPane;
        private CustomTaskPane _mainPane;
        private SettingsTaskPaneControl _settingsControl;
        private MainFormTaskPaneControl _mainControl;
        private Explorer _activeExplorer;

        public static MailItem TypeOfMailitem(MailItem mailItem)
        {
            mailItem = null;
            dynamic windowType = Globals.ThisAddIn.Application.ActiveWindow();
            if (windowType is Explorer)
            {
                // Explorer window
                Explorer explorer = windowType as Explorer;
                if (explorer.Selection.Count > 0)
                {
                    mailItem = explorer.Selection[1] as MailItem;
                }
            }
            else if (windowType is Inspector)
            {
                // Read or Compose window
                Inspector inspector = windowType as Inspector;
                mailItem = inspector.CurrentItem as MailItem;
            }

            return mailItem;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Initialize the add-in

            // Old explicit settings task pane using FormMain is no longer needed and removed
            // Setup unified task panes instead
            CreateTaskPanes();
            HookExplorerEvents();
        }

        private void CreateTaskPanes()
        {
            if (_settingsControl == null)
            {
                // Pass this add-in as ISettingsRequester so changes propagate into settingsModel
                _settingsControl = new SettingsTaskPaneControl(this);
                _settingsPane = CustomTaskPanes.Add(_settingsControl, "SaveAsPDF Settings");
                _settingsPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                _settingsPane.Width = 400;
                _settingsPane.Visible = false;
            }

            if (_mainControl == null)
            {
                _mainControl = new MainFormTaskPaneControl();
                _mainPane = CustomTaskPanes.Add(_mainControl, "SaveAsPDF");
                _mainPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                _mainPane.Width = 450;
                _mainPane.Visible = false;
            }
        }

        private void HookExplorerEvents()
        {
            _activeExplorer = Application.ActiveExplorer();
            if (_activeExplorer != null)
            {
                _activeExplorer.SelectionChange += ActiveExplorer_SelectionChange;
                // Initial load for current selection
                ActiveExplorer_SelectionChange();
            }
        }

        private void ActiveExplorer_SelectionChange()
        {
            try
            {
                if (_mainControl == null || _mainPane == null)
                    return;

                var selection = _activeExplorer?.Selection;
                MailItem mailItem = null;
                if (selection != null && selection.Count > 0)
                {
                    mailItem = selection[1] as MailItem;
                }

                // Show the main pane when there's a valid selection
                if (mailItem != null)
                {
                    _mainPane.Visible = true;
                }
                else
                {
                    _mainPane.Visible = false;
                }

                _mainControl.LoadMailItem(mailItem);
            }
            catch
            {
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (_activeExplorer != null)
            {
                _activeExplorer.SelectionChange -= ActiveExplorer_SelectionChange;
            }

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

        public void ToggleSettingsPane()
        {
            if (_settingsPane == null)
                return;
            _settingsPane.Visible = !_settingsPane.Visible;
        }

        public void ToggleMainPane()
        {
            if (_mainPane == null)
                return;
            _mainPane.Visible = !_mainPane.Visible;
        }

        // ISettingsRequester implementation so SettingsTaskPaneControl can push changes
        public void SettingsComplete(SettingsModel settings)
        {
            // Update the global settings model from the saved settings
            FormMain.settingsModel = SettingsHelpers.LoadSettingsToModel(settings);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}