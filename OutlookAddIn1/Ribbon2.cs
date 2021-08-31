using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace OutlookAddIn1
{
    public partial class MailItemRibbon
    {
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonDemo_Click(object sender, RibbonControlEventArgs e)
        {
            //get the application object 
            Outlook.Application application = Globals.ThisAddIn.Application;

            //get the active inspector object and check if is type of mailItem 
            Outlook.Inspector inspector = application.ActiveInspector();
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                MessageBox.Show("Subject: " + mailItem.Subject);

            }

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //Get the application object
            Outlook.Application application = Globals.ThisAddIn.Application;

            //Get the active Inspector object and check if it is type of MailItem
            Outlook.Inspector Inspector = application.ActiveInspector();
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                Document document = (Word.Document)Inspector.WordEditor;
                string selectedText = document.Application.Selection.Text;
                MessageBox.Show(selectedText);


            }

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //Get the application object
            Outlook.Application application = Globals.ThisAddIn.Application;

            //Get the active Inspector object and check if it is type of MailItem
            Outlook.Inspector Inspector = application.ActiveInspector();
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                //Document document = (Word.Document)Inspector.WordEditor;
                //string selectedText = document.Application.Selection.Text;

                //open form1
                Form1 frm = new Form1();
                frm.Show();



            }
        }

    }
}
