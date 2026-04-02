using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SaveAsPDF
{
    /// <summary>
    /// Represents the contacts selection form, allowing the user to select and filter employees from a list.
    /// A TreeView on the right shows all Outlook contact folders / sources.
    /// </summary>
    public partial class FormContacts : Form
    {
        /// <summary>
        /// The form that requested the employee selection.
        /// </summary>
        private readonly IEmployeeRequester callingForm;

        /// <summary>
        /// The list of employees displayed in the form.
        /// </summary>
        private List<EmployeeModel> employees = new List<EmployeeModel>();

        /// <summary>
        /// Indicates whether the dialog is in leader selection mode.
        /// </summary>
        private readonly bool _selectingLeader;

        /// <summary>
        /// Prevents re-loading the folder tree on every activation.
        /// </summary>
        private bool _activated;

        /// <summary>
        /// Timer used to animate the loading label.
        /// </summary>
        private Timer _loadingTimer;

        /// <summary>
        /// Current dot count for the loading animation.
        /// </summary>
        private int _dotCount;

        /// <summary>
        /// Initializes a new instance of the <see cref="FormContacts"/> class.
        /// </summary>
        /// <param name="caller">The form that requested the employee selection.</param>
        public FormContacts(IEmployeeRequester caller) : this(caller, false) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="FormContacts"/> class with leader selection mode.
        /// </summary>
        /// <param name="caller">The form that requested the employee selection.</param>
        /// <param name="selectingLeader">Whether the dialog is selecting a project leader.</param>
        public FormContacts(IEmployeeRequester caller, bool selectingLeader)
        {
            InitializeComponent();
            callingForm = caller;
            _selectingLeader = selectingLeader;
            if (_selectingLeader)
                dgvContacs.CellContentClick += dgvContacs_CellContentClick;
        }

        /// <summary>
        /// Handles the Load event of the form. Initializes the DataGridView and sets up column headers and visibility.
        /// </summary>
        private void FormContacts_Load(object sender, EventArgs e)
        {
            lblLoading.Visible = true;
            dgvContacs.DataSource = employees;
            ConfigureColumns();
        }

        /// <summary>
        /// Handles the Activated event of the form.
        /// On first activation, starts the loading animation and defers heavy work
        /// so the form can paint immediately.
        /// </summary>
        private void FormContacts_Activated(object sender, EventArgs e)
        {
            if (_activated)
                return;
            _activated = true;

            StartLoadingAnimation();
            BeginInvoke(new Action(LoadContacts));
        }

        /// <summary>
        /// Starts the animated loading label with Hebrew text.
        /// </summary>
        private void StartLoadingAnimation()
        {
            _dotCount = 0;
            lblLoading.Text = "טוען אנשי קשר";
            lblLoading.Visible = true;
            _loadingTimer = new Timer { Interval = 400 };
            _loadingTimer.Tick += (s, ev) =>
            {
                _dotCount = (_dotCount % 3) + 1;
                lblLoading.Text = "טוען אנשי קשר" + new string('.', _dotCount);
            };
            _loadingTimer.Start();
        }

        /// <summary>
        /// Stops the loading animation and hides the label.
        /// </summary>
        private void StopLoadingAnimation()
        {
            if (_loadingTimer != null)
            {
                _loadingTimer.Stop();
                _loadingTimer.Dispose();
                _loadingTimer = null;
            }
            lblLoading.Visible = false;
        }

        /// <summary>
        /// Performs the heavy Outlook work: builds the folder tree and loads the first folder's contacts.
        /// Called via BeginInvoke so the form is already visible.
        /// </summary>
        private void LoadContacts()
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                PopulateFolderTree();

                // Auto-select the first contact folder node to trigger loading
                TreeNode firstFolder = FindFirstFolderNode(tvFolders.Nodes);
                if (firstFolder != null)
                {
                    tvFolders.SelectedNode = firstFolder;
                }
                else
                {
                    // Fallback: load from default contacts folder
                    employees = OutlookProcessor.ListContacts();
                    BindGrid(employees);
                }
            }
            finally
            {
                Cursor.Current = Cursors.Default;
                StopLoadingAnimation();
            }
        }

        /// <summary>
        /// Populates the TreeView with Outlook contact folder hierarchy.
        /// </summary>
        private void PopulateFolderTree()
        {
            tvFolders.BeginUpdate();
            tvFolders.Nodes.Clear();

            var tree = OutlookProcessor.GetContactFolderTree();
            foreach (var store in tree)
            {
                var storeNode = new TreeNode(store.Name) { Tag = store };
                AddFolderNodes(storeNode, store.SubFolders);
                tvFolders.Nodes.Add(storeNode);
            }

            tvFolders.ExpandAll();
            tvFolders.EndUpdate();
        }

        /// <summary>
        /// Recursively adds child folder nodes to a parent tree node.
        /// </summary>
        private void AddFolderNodes(TreeNode parentNode, List<ContactFolderInfo> folders)
        {
            foreach (var folder in folders)
            {
                var node = new TreeNode(folder.Name) { Tag = folder };
                AddFolderNodes(node, folder.SubFolders);
                parentNode.Nodes.Add(node);
            }
        }

        /// <summary>
        /// Finds the first tree node that represents an actual contact folder (has an EntryID).
        /// </summary>
        private TreeNode FindFirstFolderNode(TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                var info = node.Tag as ContactFolderInfo;
                if (info != null && info.EntryID != null)
                    return node;
                var child = FindFirstFolderNode(node.Nodes);
                if (child != null)
                    return child;
            }
            return null;
        }

        /// <summary>
        /// Handles folder selection in the TreeView. Loads contacts from the selected folder.
        /// </summary>
        private void tvFolders_AfterSelect(object sender, TreeViewEventArgs e)
        {
            var info = e.Node.Tag as ContactFolderInfo;
            if (info == null)
                return;

            Cursor.Current = Cursors.WaitCursor;
            lblLoading.Text = "טוען אנשי קשר...";
            lblLoading.Visible = true;
            lblLoading.Refresh();

            try
            {
                if (info.EntryID != null)
                {
                    // Leaf contact folder – load its contacts
                    employees = OutlookProcessor.ListContactsFromFolder(info.EntryID, info.StoreID);
                }
                else
                {
                    // Store-level node – aggregate contacts from all child folders
                    employees = new List<EmployeeModel>();
                    CollectContactsFromChildren(info.SubFolders, employees);
                }

                txtFilter.Text = string.Empty;
                BindGrid(employees);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
                lblLoading.Visible = false;
            }
        }

        /// <summary>
        /// Recursively collects contacts from all child folders.
        /// </summary>
        private void CollectContactsFromChildren(List<ContactFolderInfo> folders, List<EmployeeModel> target)
        {
            foreach (var folder in folders)
            {
                if (folder.EntryID != null)
                    target.AddRange(OutlookProcessor.ListContactsFromFolder(folder.EntryID, folder.StoreID));
                CollectContactsFromChildren(folder.SubFolders, target);
            }
        }

        /// <summary>
        /// Binds the given list to the DataGridView and configures columns.
        /// </summary>
        private void BindGrid(List<EmployeeModel> list)
        {
            dgvContacs.DataSource = null;
            dgvContacs.DataSource = list;
            ConfigureColumns();
        }

        /// <summary>
        /// Handles the Click event of the OK button. Passes the selected employee to the calling form and closes this form.
        /// </summary>
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (dgvContacs.SelectedRows.Count == 0)
                return;

            var row = dgvContacs.SelectedRows[0];
            var employee = new EmployeeModel
            {
                FirstName = row.Cells[1].Value as string ?? " ",
                LastName = row.Cells[2].Value as string ?? " ",
                EmailAddress = row.Cells[3].Value as string ?? " ",
                IsLeader = _selectingLeader
            };

            callingForm.EmployeeComplete(employee);
            Close();
        }

        /// <summary>
        /// Handles the DoubleClick event of the DataGridView. Selects the employee and closes the form.
        /// </summary>
        private void dgvContacs_DoubleClick(object sender, EventArgs e)
        {
            btnOK_Click(sender, e);
        }

        /// <summary>
        /// Handles the TextChanged event of the filter textbox. Filters the employee list based on the entered text.
        /// </summary>
        private void txtFilter_TextChanged(object sender, EventArgs e)
        {
            string filter = txtFilter.Text;
            if (string.IsNullOrEmpty(filter))
            {
                BindGrid(employees);
                return;
            }

            var result = new List<EmployeeModel>();
            for (int i = 0; i < employees.Count; i++)
            {
                var emp = employees[i];
                if ((emp.EmailAddress != null && emp.EmailAddress.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    (emp.FirstName != null && emp.FirstName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0) ||
                    (emp.LastName != null && emp.LastName.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    result.Add(emp);
                }
            }
            BindGrid(result);
        }

        /// <summary>
        /// Handles the Click event of the Cancel button. Closes the form without making a selection.
        /// </summary>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ConfigureColumns()
        {
            if (dgvContacs.Columns.Count >= 5)
            {
                dgvContacs.Columns[0].Visible = false;
                dgvContacs.Columns[1].HeaderText = "שם פרטי";
                dgvContacs.Columns[2].HeaderText = "שם משפחה";
                dgvContacs.Columns[3].HeaderText = "אימייל";
                dgvContacs.Columns[4].Visible = false;
            }

            foreach (DataGridViewColumn col in dgvContacs.Columns)
            {
                if (string.Equals(col.DataPropertyName, "FullName", StringComparison.OrdinalIgnoreCase))
                    col.Visible = false;
            }

            foreach (DataGridViewColumn col in dgvContacs.Columns)
            {
                if (string.Equals(col.DataPropertyName, "IsLeader", StringComparison.OrdinalIgnoreCase))
                {
                    if (_selectingLeader)
                    {
                        col.Visible = true;
                        col.HeaderText = "מוביל";
                        col.ReadOnly = false;
                        col.DisplayIndex = 0;
                    }
                    else
                    {
                        col.Visible = false;
                    }
                    break;
                }
            }
        }

        private void dgvContacs_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            if (!string.Equals(dgvContacs.Columns[e.ColumnIndex].DataPropertyName, "IsLeader", StringComparison.OrdinalIgnoreCase)) return;

            dgvContacs.CommitEdit(DataGridViewDataErrorContexts.Commit);

            foreach (var emp in employees)
                emp.IsLeader = false;

            var currentList = dgvContacs.DataSource as List<EmployeeModel>;
            if (currentList != null && e.RowIndex < currentList.Count)
                currentList[e.RowIndex].IsLeader = true;

            dgvContacs.ClearSelection();
            dgvContacs.Rows[e.RowIndex].Selected = true;
            dgvContacs.Invalidate();
        }
    }
}
