using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    /// <summary>
    /// A lightweight Explorer-like address bar with breadcrumb display and edit mode.
    /// </summary>
    public class ExplorerAddressBar : UserControl
    {
        private readonly FlowLayoutPanel _breadcrumbPanel;
        private readonly TextBox _editBox;
        private bool _editMode;
        private string _currentPath = string.Empty;

        public event EventHandler<PathChangedEventArgs> PathChanged;
        public event EventHandler<PathConfirmedEventArgs> PathConfirmed;

        public string Path
        {
            get => _currentPath;
            set
            {
                var normalized = NormalizePath(value);
                if (!string.Equals(_currentPath, normalized, StringComparison.OrdinalIgnoreCase))
                {
                    _currentPath = normalized;
                    RefreshBreadcrumbs();
                    _editBox.Text = _currentPath;
                    PathChanged?.Invoke(this, new PathChangedEventArgs(_currentPath));
                }
            }
        }

        public ExplorerAddressBar()
        {
            Height = 24;
            BackColor = SystemColors.Window;
            BorderStyle = BorderStyle.FixedSingle;
            Padding = Padding.Empty;

            _breadcrumbPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = false,
                Padding = Padding.Empty,
                Margin = Padding.Empty,
                BackColor = SystemColors.Window
            };
            _breadcrumbPanel.Click += (s, e) => EnterEditMode();

            _editBox = new TextBox
            {
                Visible = false,
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.WindowText
            };
            _editBox.KeyDown += EditBox_KeyDown;
            _editBox.Leave += (s, e) => ExitEditMode(applyChanges: true);

            Controls.Add(_breadcrumbPanel);
            Controls.Add(_editBox);
        }

        private void EditBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                ExitEditMode(applyChanges: true, confirm: true);
            }
            else if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true;
                ExitEditMode(applyChanges: false);
            }
        }

        private void EnterEditMode()
        {
            if (_editMode)
                return;

            _editMode = true;
            _editBox.Text = _currentPath;
            _editBox.Visible = true;
            _breadcrumbPanel.Visible = false;
            _editBox.Focus();
            _editBox.SelectAll();
        }

        private void ExitEditMode(bool applyChanges, bool confirm = false)
        {
            if (!_editMode)
                return;

            _editMode = false;
            if (applyChanges)
            {
                var entered = _editBox.Text?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(entered))
                {
                    Path = entered;
                    if (confirm)
                        PathConfirmed?.Invoke(this, new PathConfirmedEventArgs(_currentPath));
                }
            }
            _editBox.Visible = false;
            _breadcrumbPanel.Visible = true;
        }

        private void RefreshBreadcrumbs()
        {
            _breadcrumbPanel.Controls.Clear();

            if (string.IsNullOrEmpty(_currentPath))
                return;

            var segments = SplitPathSegments(_currentPath);
            var cumulative = new List<string>();

            for (int i = 0; i < segments.Count; i++)
            {
                cumulative.Add(segments[i]);
                string pathUpToHere = string.Join("\\", cumulative);

                if (i > 0)
                {
                    var sep = new Label
                    {
                        Text = "›",
                        AutoSize = true,
                        Margin = Padding.Empty,
                        ForeColor = SystemColors.GrayText
                    };
                    _breadcrumbPanel.Controls.Add(sep);
                }

                var lbl = new Label
                {
                    Text = segments[i],
                    AutoSize = true,
                    Margin = Padding.Empty,
                    Padding = Padding.Empty,
                    BackColor = SystemColors.Window,
                    ForeColor = SystemColors.WindowText,
                    Cursor = Cursors.Hand,
                    Tag = pathUpToHere
                };
                lbl.Click += (s, e) => OnSegmentClicked(pathUpToHere, lbl);
                _breadcrumbPanel.Controls.Add(lbl);
            }
        }

        private void OnSegmentClicked(string path, Control anchor)
        {
            Path = path;

            var menu = BuildFolderMenu(path);
            if (menu.Items.Count > 0)
            {
                menu.Show(anchor, new Point(0, anchor.Height));
            }
            else
            {
                PathConfirmed?.Invoke(this, new PathConfirmedEventArgs(path));
            }
        }

        private ContextMenuStrip BuildFolderMenu(string path)
        {
            var menu = new ContextMenuStrip
            {
                ShowImageMargin = false,
                AutoClose = true,
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.WindowText
            };

            string basePath = NormalizeForDirectoryExists(path);
            if (Directory.Exists(basePath))
            {
                try
                {
                    var dirs = Directory.GetDirectories(basePath)
                        .Select(d => new DirectoryInfo(d))
                        .OrderBy(d => d.Name)
                        .ToList();

                    foreach (var dir in dirs)
                    {
                        var item = new ToolStripMenuItem(dir.Name)
                        {
                            Tag = dir.FullName
                        };
                        item.Click += (s, e) =>
                        {
                            Path = dir.FullName;
                            PathConfirmed?.Invoke(this, new PathConfirmedEventArgs(dir.FullName));
                        };
                        menu.Items.Add(item);
                    }
                }
                catch
                {
                    // Ignore directory enumeration errors
                }
            }

            return menu;
        }

        private static List<string> SplitPathSegments(string path)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(path))
                return list;

            // Normalize separators
            var parts = path.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
                list.Add(part);
            return list;
        }

        private static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return string.Empty;

            // Ensure drive roots are treated as roots (avoid current-directory resolution)
            if (path.Length == 2 && path[1] == ':')
                path = path + "\\";

            try
            {
                return System.IO.Path.GetFullPath(path.Trim());
            }
            catch
            {
                return path.Trim();
            }
        }

        private static string NormalizeForDirectoryExists(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return string.Empty;

            if (path.Length == 2 && path[1] == ':')
                return path + "\\";

            return path;
        }
    }

    public class PathChangedEventArgs : EventArgs
    {
        public string Path { get; }
        public PathChangedEventArgs(string path) => Path = path;
    }

    public class PathConfirmedEventArgs : EventArgs
    {
        public string Path { get; }
        public PathConfirmedEventArgs(string path) => Path = path;
    }
}
