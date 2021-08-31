using Microsoft.Office.Tools.Ribbon;

namespace SaveAsPDF1
{
    public partial class ExlorerRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //open Main
            frmMain frm = new frmMain();
            frm.Show();
        }
    }
}
