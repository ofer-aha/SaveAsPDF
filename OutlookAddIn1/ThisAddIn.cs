using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SaveAsPDF1
{
    public static class PGlobals
    {
        private static string _ROOTDRIVE;
        public static string ROOTDRIVE
        {
            get
            {
                return _ROOTDRIVE;
            }
            set => _ROOTDRIVE = @"J:\";

        }
    }
    public partial class ThisAddIn
    {
        Outlook.Explorer currentExplorer; // = null;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Get the application object 
            Outlook.Application application = this.Application;

            //Get the inspector object 
            Outlook.Inspectors inspectors = application.Inspectors;

            //Get the avtive inspector object 
            Outlook.Inspector activeInspector = application.ActiveInspector();
         
            currentExplorer = this.Application.ActiveExplorer();
         



            // if (activeInspector != null)
            // {
            //Get the title of the active iten when the outlook start. 
            //     MessageBox.Show("Active inspector: " + activeInspector.Caption);

            // }
            //Get the explorer object 
           // Outlook.Explorers explorers = application.Explorers;

            //Get the ACTIVE explorer object 
           // Outlook.Explorer activeExplorer = application.ActiveExplorer();
           // if (activeExplorer != null)
           // {
                //Get the title of the active folder when the outlook start 
           //     MessageBox.Show("Active Explorer: " + activeExplorer.Caption);

           // }

            //Add a new Inspector to the application 
          //  inspectors.NewInspector +=
          //      new Outlook.InspectorsEvents_NewInspectorEventHandler(
          //          Inspectors_AddTextToNewMail);

          //  inspectors.NewInspector +=
          //      new Outlook.InspectorsEvents_NewInspectorEventHandler(
          //          Inspectors_RegisterEventWordDocument);


            //Subscribe to the ItemSend event, that is triggered when an email is sent
          //  application.ItemSend +=
          //      new Outlook.ApplicationEvents_11_ItemSendEventHandler(
           //     ItemSend_BeforSend);

        }



        private void CurrentExplorer_Event()
        {
            Outlook.MAPIFolder selectedFolder =
                this.Application.ActiveExplorer().CurrentFolder;
            String expMessage = "Your current folder is "
                + selectedFolder.Name + ".\n";
            String itemMessage = "Item is unknown.";
            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];
                    if (selObject is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem =
                            (selObject as Outlook.MailItem);
                        itemMessage = "The item is an e-mail message." +
                            " The subject is " + mailItem.Subject + ".";
                        mailItem.Display(false);
                    }
                 
                }
                expMessage += itemMessage;
            }
            catch (System.Exception ex)
            {
                expMessage = ex.Message;
            }
            MessageBox.Show(expMessage);
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

        #endregion
    }
}
