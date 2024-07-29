using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace SaveAsPDF
{
    public partial class FormContacts : Form
    {
        private readonly IEmployeeRequester callingForm;
        List<EmployeeModel> employees = new List<EmployeeModel>();



        public FormContacts(IEmployeeRequester caller)
        {
            InitializeComponent();

            callingForm = caller;
        }

        private void FormContacts_Load(object sender, EventArgs e)
        {
            lblLoading.Visible = true;
            dgvContacs.DataSource = employees;
            dgvContacs.Columns[0].Visible = false;
            dgvContacs.Columns[1].HeaderText = "שם פרטי";
            dgvContacs.Columns[2].HeaderText = "שם משפחה";
            dgvContacs.Columns[3].HeaderText = "אימייל";
            dgvContacs.Columns[4].Visible = false;
        }
        private void FormContacts_Activated(object sender, EventArgs e)
        {
            //TODO: is the right event to do this? 

            // Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor;

            employees = OutlookProcessor.ListContacts();
            dgvContacs.DataSource = employees;

            // Set cursor as default arrow
            Cursor.Current = Cursors.Default;
            lblLoading.Visible = false;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {

            EmployeeModel employee = new EmployeeModel();
            if (dgvContacs.SelectedRows[0].Cells[1].Value == null)
            {
                employee.FirstName = " ";
            }
            else
            {
                employee.FirstName = dgvContacs.SelectedRows[0].Cells[1].Value.ToString();
            }

            if (dgvContacs.SelectedRows[0].Cells[2].Value == null)
            {
                employee.LastName = " ";
            }
            else
            {
                employee.LastName = dgvContacs.SelectedRows[0].Cells[2].Value.ToString();
            }

            if (dgvContacs.SelectedRows[0].Cells[3].Value == null)
            {
                employee.EmailAddress = " ";
            }
            else
            {
                employee.EmailAddress = dgvContacs.SelectedRows[0].Cells[3].Value.ToString();
            }

            callingForm.EmployeeComplete(employee);

            Close();

        }


        private void dgvContacs_DoubleClick(object sender, EventArgs e)
        {
            btnOK_Click(sender, e);


        }

        private void txtFilter_TextChanged(object sender, EventArgs e)
        {

            var result = employees.Where(x => x.EmailAddress.Contains(txtFilter.Text) ||
                                                x.FirstName.Contains(txtFilter.Text) ||
                                                x.FirstName.Contains(txtFilter.Text)).ToList();
            dgvContacs.DataSource = result;

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
