using Microsoft.Office.Tools.Ribbon;

namespace SaveAsPDF
{
    public partial class SaveAsPDFRibbon
    {
        private void SaveAsPDFRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonDemo_Click(object sender, RibbonControlEventArgs e)
        {
            frmMain frm = new frmMain();
            frm.ShowDialog();
        }


    }
}
