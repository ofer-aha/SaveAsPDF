using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace SaveAsPDF
{
    public partial class  SaveAsPDFRibbon
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
