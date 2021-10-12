using System.Windows.Forms;

namespace SaveAsPDF
{
    public partial class frmSettings : Form
    {
        public frmSettings()
        {
            InitializeComponent();
        }

        private void bntCancel_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }
    }
}