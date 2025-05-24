using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SaveAsPDF
{
    /// <summary>
    /// Represents the contacts selection form, allowing the user to select and filter employees from a list.
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
        /// Initializes a new instance of the <see cref="FormContacts"/> class.
        /// </summary>
        /// <param name="caller">The form that requested the employee selection.</param>
        public FormContacts(IEmployeeRequester caller)
        {
            InitializeComponent();
            callingForm = caller;
        }

        /// <summary>
        /// Handles the Load event of the form. Initializes the DataGridView and sets up column headers and visibility.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event arguments.</param>
        private void FormContacts_Load(object sender, EventArgs e)
        {
            lblLoading.Visible = true;
            dgvContacs.DataSource = employees;
            if (dgvContacs.Columns.Count >= 5)
            {
                dgvContacs.Columns[0].Visible = false;
                dgvContacs.Columns[1].HeaderText = "שם פרטי";
                dgvContacs.Columns[2].HeaderText = "שם משפחה";
                dgvContacs.Columns[3].HeaderText = "אימייל";
                dgvContacs.Columns[4].Visible = false;
            }
        }

        /// <summary>
        /// Handles the Activated event of the form. Loads the contacts from Outlook and updates the DataGridView.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event arguments.</param>
        private void FormContacts_Activated(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            employees = OutlookProcessor.ListContacts();
            dgvContacs.DataSource = null;
            dgvContacs.DataSource = employees;

            Cursor.Current = Cursors.Default;
            lblLoading.Visible = false;
        }

        /// <summary>
        /// Handles the Click event of the OK button. Passes the selected employee to the calling form and closes this form.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event arguments.</param>
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (dgvContacs.SelectedRows.Count == 0)
                return;

            var row = dgvContacs.SelectedRows[0];
            var employee = new EmployeeModel
            {
                FirstName = row.Cells[1].Value as string ?? " ",
                LastName = row.Cells[2].Value as string ?? " ",
                EmailAddress = row.Cells[3].Value as string ?? " "
            };

            callingForm.EmployeeComplete(employee);
            Close();
        }

        /// <summary>
        /// Handles the DoubleClick event of the DataGridView. Selects the employee and closes the form.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event arguments.</param>
        private void dgvContacs_DoubleClick(object sender, EventArgs e)
        {
            btnOK_Click(sender, e);
        }

        /// <summary>
        /// Handles the TextChanged event of the filter textbox. Filters the employee list based on the entered text.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event arguments.</param>
        private void txtFilter_TextChanged(object sender, EventArgs e)
        {
            string filter = txtFilter.Text;
            if (string.IsNullOrEmpty(filter))
            {
                dgvContacs.DataSource = employees;
                return;
            }

            var result = new List<EmployeeModel>();
            for (int i = 0; i < employees.Count; i++)
            {
                var emp = employees[i];
                if ((emp.EmailAddress != null && emp.EmailAddress.Contains(filter)) ||
                    (emp.FirstName != null && emp.FirstName.Contains(filter)) ||
                    (emp.LastName != null && emp.LastName.Contains(filter)))
                {
                    result.Add(emp);
                }
            }
            dgvContacs.DataSource = result;
        }

        /// <summary>
        /// Handles the Click event of the Cancel button. Closes the form without making a selection.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event arguments.</param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
