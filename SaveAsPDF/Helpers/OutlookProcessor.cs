using Microsoft.Office.Interop.Outlook;
using SaveAsPDF.Models;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using Exception = System.Exception;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SaveAsPDF.Helpers
{
    public class OutlookProcessor
    {

        public static List<EmployeeModel> ListContacts()
        {
            List<EmployeeModel> output = new List<EmployeeModel>();

            Outlook.Application app = new Outlook.Application();
            NameSpace NameSpace = app.GetNamespace("MAPI");
            MAPIFolder ContactsFolder = NameSpace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            Items ContactItems = ContactsFolder.Items;
            try
            {
                foreach (ContactItem item in ContactItems)
                {
                    if (!string.IsNullOrEmpty(item.Email1Address))
                    {
                        EmployeeModel employee = new EmployeeModel();
                        employee.EmailAddress = item.Email1Address;
                        employee.FirstName = item.FirstName;
                        employee.LastName = item.LastName;

                        output.Add(employee);
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show(ex.Message, "SaveAsPDF");
            }
            return output;

        }



        public void EnumerateFoldersInDefaultStore()
        {
            Folder root = Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder() as Folder;
            EnumerateFolders(root);
        }

        /// <summary>
        /// Uses recursion to enumerate Outlook sub-folders. 
        /// </summary>
        /// <param name="folder"></param>
        public void EnumerateFolders(Folder folder)
        {
            Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Folder childFolder in childFolders)
                {
                    // Write the folder path.
                    Debug.WriteLine(childFolder.FolderPath);
                    // Call EnumerateFolders using childFolder.
                    EnumerateFolders(childFolder);
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="findLastName"></param>
        public static void AccessContacts(string findLastName)
        {
            MAPIFolder folderContacts = Globals.ThisAddIn.Application.ActiveExplorer().Session.
                GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            Items searchFolder = folderContacts.Items;
            int counter = 0;
            foreach (ContactItem foundContact in searchFolder)
            {
                if (foundContact.LastName.Contains(findLastName))
                {
                    foundContact.Display(false);
                    counter = counter + 1;
                }
            }
            MessageBox.Show($"You have {counter} contacts with last names that contain {findLastName}.");
        }

        public static void FindContact(string inString)
        {
            NameSpace outlookNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            MAPIFolder contactsFolder = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

            Items contactItems = contactsFolder.Items;

            try
            {
                //Outlook.ContactItem contact = (Outlook.ContactItem)contactItems.Find(String.Format($"[FirstName]='{firstName.Trim()}'"));
                ContactItem contact = (ContactItem)contactItems.Find(inString.Trim());
                if (contact != null)
                {
                    contact.Display(true);
                }
                else
                {
                    MessageBox.Show("The contact information was not found.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


    }
}
