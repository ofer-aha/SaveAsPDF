using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SaveAsPDF
{
    /// <summary>
    /// COM-visible host control for <see cref="SettingsTaskPaneControl"/>.
    /// Required because <see cref="SettingsTaskPaneControl"/> needs an
    /// <see cref="ISettingsRequester"/> in its constructor, whereas
    /// <c>ICTPFactory.CreateCTP</c> requires a public parameterless constructor.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("6E4B2E1A-3F9D-4C5E-A7B8-2D1C0F3E5A9B")]
    [ProgId("SaveAsPDF.SettingsTaskPaneHost")]
    public class SettingsTaskPaneHost : UserControl
    {
        private SettingsTaskPaneControl _inner;

        public SettingsTaskPaneHost()
        {
            Dock = DockStyle.Fill;
            RightToLeft = RightToLeft.Yes;
        }

        /// <summary>
        /// Creates the inner <see cref="SettingsTaskPaneControl"/> and wires it to
        /// the provided <paramref name="caller"/>.
        /// Call this once after the CTP factory has created the host.
        /// </summary>
        public void Initialize(ISettingsRequester caller)
        {
            if (caller == null) throw new ArgumentNullException(nameof(caller));
            if (_inner != null) return; // already initialized

            _inner = new SettingsTaskPaneControl(caller);
            _inner.Dock = DockStyle.Fill;
            Controls.Add(_inner);
        }

        /// <summary>
        /// Returns the inner settings control, or <c>null</c> if not yet initialized.
        /// </summary>
        public SettingsTaskPaneControl InnerControl => _inner;
    }
}
