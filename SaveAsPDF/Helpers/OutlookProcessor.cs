using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

using SaveAsPDF.Models;
using System.Windows.Forms;
using Exception = System.Exception;

namespace SaveAsPDF.Helpers
{
    public class OutlookProcessor
    {

        public static List<EmployeeModel> ListContacts()
        {
            List<EmployeeModel> output = new List<EmployeeModel>();
            
            Outlook.Application app = new Outlook.Application();
            Outlook.NameSpace NameSpace = app.GetNamespace("MAPI");
            Outlook.MAPIFolder ContactsFolder = NameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            Outlook.Items ContactItems = ContactsFolder.Items;
            try
            {
                foreach (Outlook.ContactItem item in ContactItems)
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
                MessageBox.Show(ex.Message,"SaveAsPDF");
            }
            return output;
            
        }



        public void EnumerateFoldersInDefaultStore()
        {
            Outlook.Folder root = Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;
            EnumerateFolders(root);
        }

        // Uses recursion to enumerate Outlook subfolders.
        public void EnumerateFolders(Outlook.Folder folder)
        {
            Outlook.Folders childFolders =
                folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Outlook.Folder childFolder in childFolders)
                {
                    // Write the folder path.
                    Debug.WriteLine(childFolder.FolderPath);
                    // Call EnumerateFolders using childFolder.
                    EnumerateFolders(childFolder);
                }
            }
        }



        public static void AccessContacts(string findLastName)
        {
            Outlook.MAPIFolder folderContacts = Globals.ThisAddIn.Application.ActiveExplorer().Session.
                GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            Outlook.Items searchFolder = folderContacts.Items;
            int counter = 0;
            foreach (Outlook.ContactItem foundContact in searchFolder)
            {
                if (foundContact.LastName.Contains(findLastName))
                {
                    foundContact.Display(false);
                    counter = counter + 1;
                }
            }
            MessageBox.Show("You have " + counter +
                " contacts with last names that contain "
                + findLastName + ".");
        }

        public static void FindContact(string inString)
        {
            Outlook.NameSpace outlookNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            Outlook.MAPIFolder contactsFolder = outlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

            Outlook.Items contactItems = contactsFolder.Items;

            try
            {
                //Outlook.ContactItem contact = (Outlook.ContactItem)contactItems.Find(String.Format($"[FirstName]='{firstName.Trim()}'"));
                Outlook.ContactItem contact = (Outlook.ContactItem)contactItems.Find(inString.Trim());
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
