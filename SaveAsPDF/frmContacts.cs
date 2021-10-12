using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;

namespace SaveAsPDF
{
    public partial class frmContacts : Form
    {
        private IEmployeeRequester callingForm;
        public frmContacts(IEmployeeRequester caller)
        {
            InitializeComponent();

            callingForm = caller; 
        }

        List<EmployeeModel> employees = new List<EmployeeModel>();

        private void frmContacts_Load(object sender, EventArgs e)
        {
            
            employees = OutlookProcessor.ListContacts();
            
            dgvContacs.DataSource = employees;
            
            dgvContacs.Columns[0].Visible = false;
            dgvContacs.Columns[1].HeaderText = "שם פרטי";
            dgvContacs.Columns[2].HeaderText = "שם מפשחה";
            dgvContacs.Columns[3].HeaderText = "אימייל";
            dgvContacs.Columns[4].Visible = false;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            EmployeeModel employee = new EmployeeModel
            {
                FirstName = dgvContacs.SelectedRows[0].Cells[1].Value.ToString(),
                LastName = dgvContacs.SelectedRows[0].Cells[2].Value.ToString(),
                EmailAddress = dgvContacs.SelectedRows[0].Cells[3].Value.ToString()
            };



            callingForm.EmployeeComplete(employee);
            
            this.Close();

        }


        private void dgvContacs_DoubleClick(object sender, EventArgs e)
        {


            EmployeeModel employee = new EmployeeModel
            {
                FirstName = dgvContacs.SelectedRows[0].Cells[1].Value.ToString(),
                LastName = dgvContacs.SelectedRows[0].Cells[2].Value.ToString(),
                EmailAddress = dgvContacs.SelectedRows[0].Cells[3].Value.ToString()
            };


            callingForm.EmployeeComplete(employee);

            this.Close();

        }

        private void txtFilter_TextChanged(object sender, EventArgs e)
        {
            var result = employees.Where(x => x.EmailAddress.Contains(txtFilter.Text) || 
                                                x.FirstName.Contains(txtFilter.Text) || 
                                                x.FirstName.Contains(txtFilter.Text)).ToList();
            dgvContacs.DataSource = result;
        }
    }
}
