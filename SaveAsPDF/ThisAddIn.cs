using SaveAsPDF.Helpers;
using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SaveAsPDF
{
   
    
    public partial class ThisAddIn
    {
    
        public static Outlook.MailItem TypeOfMailitem(Outlook.MailItem mailItem)
        {
            mailItem = null;
            dynamic windowType = Globals.ThisAddIn.Application.ActiveWindow();
            if (windowType is Outlook.Explorer)
            {
                // frmMain Explorer
                Outlook.Explorer explorer = windowType as Outlook.Explorer;
                mailItem = explorer.Selection[1] as Outlook.MailItem;
            }
            else if (windowType is Outlook.Inspector)
            {
                // Read or Compose
                Outlook.Inspector inspector = windowType as Outlook.Inspector;
                mailItem = inspector.CurrentItem as Outlook.MailItem;
            }

            return mailItem;
        }
        

        //public static void AccessContacts(string findLastName)
        //{
        //    Outlook.MAPIFolder folderContacts = Application.ActiveExplorer().Session.GetDefaultFolder
        //                                    (Outlook.OlDefaultFolders.olFolderInbox);
        //    Outlook.Items searchFolder = folderContacts.Items;
        //    int counter = 0;
        //    foreach (Outlook.ContactItem foundContact in searchFolder)
        //    {
        //        if (foundContact.LastName.Contains(findLastName))
        //        {
        //            foundContact.Display(false);
        //            counter = counter + 1;
        //        }
        //    }
        //    MessageBox.Show("You have " + counter +
        //        " contacts with last names that contain "
        //        + findLastName + ".");
        //}

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            //this.Application.NewMail += new Microsoft.Office.Interop.Outlook.
            //                              ApplicationEvents_11_NewMailEventHandler(ThisAddIn_NewMail);

            //OutlookProcessor.FindContactEmailByName("טלי");
        }

  
        private void ThisAddIn_NewMail()
        {
            //Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)this.Application.
            //    ActiveExplorer().Session.GetDefaultFolder
            //    (Outlook.OlDefaultFolders.olFolderInbox);
            //Outlook.Items items = (Outlook.Items)inBox.Items;
            //Outlook.MailItem moveMail = null;
            //items.Restrict("[UnRead] = true");
            //Outlook.MAPIFolder destFolder = inBox.Folders["Test"];
            //foreach (object eMail in items)
            //{
            //    try
            //    {
            //        moveMail = eMail as Outlook.MailItem;
            //        if (moveMail != null)
            //        {
            //            string titleSubject = (string)moveMail.Subject;
            //            if (titleSubject.IndexOf("Test") > 0)
            //            {
            //                moveMail.Move(destFolder);
            //            }
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message);
            //    }
            //}
        }

  

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
            {
                // Note: Outlook no longer raises this event. If you have code that
                //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
            }

            #region VSTO generated code

            /// <summary>
            /// Required method for Designer support - do not modify
            /// the contents of this method with the code editor.
            /// </summary>
            private void InternalStartup()
            {
                this.Startup += new System.EventHandler(ThisAddIn_Startup);
                this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            }

            #endregion VSTO generated code
        
    }
}