using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using Exception = System.Exception;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SaveAsPDF
{
    /// <summary>
    /// IDispatch interface that exposes only the Ribbon XML callback methods.
    /// Office uses QueryInterface → IDispatch → Invoke to late-bind to these
    /// callbacks at runtime.  Keeping this as the <see cref="ComDefaultInterfaceAttribute"/>
    /// on <see cref="Connect"/> lets Office find the methods while avoiding
    /// type-library export errors from non-COM-safe members.
    /// </summary>
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("B5D8F2A1-7C3E-4A9D-8E6F-1B2C3D4E5F60")]
    public interface IRibbonCallbacks
    {
        void Ribbon_OnLoad(IRibbonUI ribbonUI);
        stdole.IPictureDisp GetButtonImage(IRibbonControl control);
        void OnToggleMainPane(IRibbonControl control);
    }

    /// <summary>
    /// COM Add-in entry point for the SaveAsPDF Outlook add-in.
    /// Implements the standard Office extensibility interfaces so that
    /// Outlook loads the add-in without requiring the VSTO runtime.
    /// </summary>
    [ComVisible(true)]
    [Guid("A36818D8-FD7B-44D1-9F18-4E2F62AB4A6E")]
    [ProgId("SaveAsPDF.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRibbonCallbacks))]
    public class Connect : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer, ISettingsRequester, IRibbonCallbacks
    {
        #region Fields

        private Application _application;
        private ICTPFactory _ctpFactory;
        private IRibbonUI _ribbon;

        private CustomTaskPane _settingsPane;
        private CustomTaskPane _mainPane;
        private SettingsTaskPaneHost _settingsHost;
        private MainFormTaskPaneControl _mainControl;
        private Explorer _activeExplorer;
        private bool _startupCompleted;

        #endregion

        #region Properties

        /// <summary>
        /// The host Outlook Application object.
        /// </summary>
        public Application Application => _application;

        #endregion

        #region IDTExtensibility2

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            _application = (Application)application;
            AddinGlobals.Connect = this;

            // For non-startup loads (e.g. enabled via COM Add-ins dialog),
            // hook explorer events now; task panes will be created once
            // CTPFactoryAvailable fires.
            if (connectMode != ext_ConnectMode.ext_cm_Startup)
            {
                _startupCompleted = true;
                System.Windows.Forms.Application.EnableVisualStyles();
                HookExplorerEvents();
            }
        }

        public void OnStartupComplete(ref Array custom)
        {
            System.Windows.Forms.Application.EnableVisualStyles();
            _startupCompleted = true;
            CreateTaskPanes();
            HookExplorerEvents();
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            UnhookExplorerEvents();
            CleanupResources();
            _application = null;
            AddinGlobals.Connect = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        #endregion

        #region ICustomTaskPaneConsumer

        public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst)
        {
            _ctpFactory = CTPFactoryInst;

            // If OnStartupComplete already ran (non-startup connect mode),
            // create task panes now that the factory is available.
            if (_startupCompleted)
            {
                CreateTaskPanes();
            }
        }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            switch (ribbonID)
            {
                case "Microsoft.Outlook.Explorer":
                case "Microsoft.Outlook.Mail.Read":
                    return GetResourceText("SaveAsPDF.Resources.customUI14.xml");
                default:
                    return null;
            }
        }

        #endregion

        #region Ribbon Callbacks

        /// <summary>Called by Outlook when the ribbon UI loads.</summary>
        public void Ribbon_OnLoad(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        /// <summary>
        /// Returns the custom button image as <c>stdole.IPictureDisp</c>.
        /// Office invokes this via IDispatch; returning a raw <see cref="Bitmap"/>
        /// works only in some Office builds, whereas <c>IPictureDisp</c> is
        /// universally supported.
        /// </summary>
        public stdole.IPictureDisp GetButtonImage(IRibbonControl control)
        {
            try
            {
                return PictureConverter.BitmapToIPictureDisp(Properties.Resources.SaveAsPDF32);
            }
            catch { }
            return null;
        }

        /// <summary>Toggles the main task pane visibility.</summary>
        public void OnToggleMainPane(IRibbonControl control)
        {
            ToggleMainPane();
        }

        #endregion

        #region Task-pane management

        private void CreateTaskPanes()
        {
            if (_ctpFactory == null || _mainPane != null) return;

            try
            {
                // --- Settings pane (via COM-visible wrapper) ---
                _settingsPane = _ctpFactory.CreateCTP(
                    "SaveAsPDF.SettingsTaskPaneHost",
                    "SaveAsPDF Settings",
                    Type.Missing);
                _settingsPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _settingsPane.Width = 400;
                _settingsPane.Visible = false;

                _settingsHost = (SettingsTaskPaneHost)_settingsPane.ContentControl;
                _settingsHost.Initialize(this);

                // --- Main pane ---
                _mainPane = _ctpFactory.CreateCTP(
                    "SaveAsPDF.MainFormTaskPaneControl",
                    "SaveAsPDF",
                    Type.Missing);
                _mainPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _mainPane.Width = 450;
                _mainPane.Visible = false;

                _mainControl = (MainFormTaskPaneControl)_mainPane.ContentControl;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException("Connect.CreateTaskPanes", ex);
                System.Windows.Forms.MessageBox.Show(
                    "SaveAsPDF: Failed to create task panes.\n" + ex.ToString(),
                    "SaveAsPDF Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void ToggleSettingsPane()
        {
            if (_settingsPane == null) return;
            _settingsPane.Visible = !_settingsPane.Visible;
        }

        public void ToggleMainPane()
        {
            // Attempt on-demand creation if task panes were not built yet
            if (_mainPane == null && _ctpFactory != null)
            {
                CreateTaskPanes();
            }

            if (_mainPane == null || _mainControl == null)
            {
                System.Windows.Forms.MessageBox.Show(
                    "SaveAsPDF task pane is not available.\n" +
                    "CTP Factory: " + (_ctpFactory != null ? "OK" : "NULL") + "\n" +
                    "Main Pane: " + (_mainPane != null ? "OK" : "NULL") + "\n" +
                    "Main Control: " + (_mainControl != null ? "OK" : "NULL"),
                    "SaveAsPDF Diagnostic",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }

            _mainPane.Visible = MainPaneBehavior.ToggleVisibility(_mainPane.Visible);

            if (MainPaneBehavior.ShouldLoadMailItemAfterToggle(_mainPane.Visible))
            {
                var explorer = _activeExplorer ?? _application.ActiveExplorer();
                var selection = explorer?.Selection;
                MailItem mailItem = null;
                if (selection != null && selection.Count > 0)
                    mailItem = selection[1] as MailItem;

                _mainControl.LoadMailItem(mailItem);
            }
        }

        #endregion

        #region Explorer events

        private void HookExplorerEvents()
        {
            _activeExplorer = _application.ActiveExplorer();
            if (_activeExplorer != null)
                _activeExplorer.SelectionChange += ActiveExplorer_SelectionChange;
        }

        private void UnhookExplorerEvents()
        {
            if (_activeExplorer != null)
            {
                _activeExplorer.SelectionChange -= ActiveExplorer_SelectionChange;
                _activeExplorer = null;
            }
        }

        private void ActiveExplorer_SelectionChange()
        {
            try
            {
                if (!MainPaneBehavior.ShouldRefreshOnSelectionChange(
                        _mainControl != null,
                        _mainPane != null,
                        _mainPane != null && _mainPane.Visible))
                    return;

                var selection = _activeExplorer?.Selection;
                MailItem mailItem = null;
                if (selection != null && selection.Count > 0)
                    mailItem = selection[1] as MailItem;

                _mainControl.LoadMailItem(mailItem);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException("Connect.ActiveExplorer_SelectionChange", ex);
            }
        }

        #endregion

        #region ISettingsRequester

        public void SettingsComplete(SettingsModel settings)
        {
            MainFormTaskPaneControl.settingsModel = SettingsHelpers.LoadSettingsToModel(settings);
        }

        #endregion

        #region Static helpers

        /// <summary>
        /// Determines the active <see cref="MailItem"/> from the current Outlook window.
        /// </summary>
        public static MailItem TypeOfMailitem(MailItem mailItem)
        {
            mailItem = null;
            dynamic windowType = AddinGlobals.OutlookApp.ActiveWindow();
            if (windowType is Explorer explorer)
            {
                if (explorer.Selection.Count > 0)
                    mailItem = explorer.Selection[1] as MailItem;
            }
            else if (windowType is Inspector inspector)
            {
                mailItem = inspector.CurrentItem as MailItem;
            }
            return mailItem;
        }

        #endregion

        #region Cleanup

        private void CleanupResources()
        {
            try
            {
                OfficeHelpers.ReleaseWordInstance();
                ComboBoxExtensions.ClearDirectoryCache();

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException("Connect.CleanupResources", ex);
            }
        }

        #endregion

        #region COM Registration

        private const string _addinKeyPath = @"Software\Microsoft\Office\Outlook\Addins\SaveAsPDF.Connect";

        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(_addinKeyPath))
            {
                key.SetValue("FriendlyName", "SaveAsPDF");
                key.SetValue("Description", "Convert Outlook mail messages to PDF file");
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord); // Load at startup
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            Registry.CurrentUser.DeleteSubKey(_addinKeyPath, throwOnMissingSubKey: false);
        }

        #endregion

        #region Utilities

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null) return string.Empty;
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        #endregion
    }

    /// <summary>
    /// Converts a <see cref="Image"/> to <c>stdole.IPictureDisp</c> by
    /// exposing the <c>protected static</c> method
    /// <see cref="System.Windows.Forms.AxHost.GetIPictureDispFromPicture"/>.
    /// This is the standard .NET Framework pattern for Ribbon XML image callbacks.
    /// </summary>
    [ComVisible(false)]
    internal class PictureConverter : System.Windows.Forms.AxHost
    {
        private PictureConverter() : base(string.Empty) { }

        public static stdole.IPictureDisp BitmapToIPictureDisp(Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }
}
