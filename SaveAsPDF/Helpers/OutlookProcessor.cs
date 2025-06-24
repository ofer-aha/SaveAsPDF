using Microsoft.Office.Interop.Outlook;
using SaveAsPDF.Models;
using System.Collections.Generic;
using System.Diagnostics;
using Exception = System.Exception;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SaveAsPDF.Helpers
{
    public class OutlookProcessor
    {
        public static List<EmployeeModel> ListContacts()
        {
            var output = new List<EmployeeModel>();

            Outlook.Application app = null;
            NameSpace nameSpace = null;
            MAPIFolder contactsFolder = null;
            Items contactItems = null;

            try
            {
                app = new Outlook.Application();
                nameSpace = app.GetNamespace("MAPI");
                contactsFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                contactItems = contactsFolder.Items;

                int count = contactItems.Count;
                for (int i = 1; i <= count; i++)
                {
                    var item = contactItems[i] as ContactItem;
                    if (item != null && !string.IsNullOrEmpty(item.Email1Address))
                    {
                        var employee = new EmployeeModel
                        {
                            EmailAddress = item.Email1Address,
                            FirstName = item.FirstName,
                            LastName = item.LastName
                        };
                        output.Add(employee);
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                XMessageBox.Show(
                    ex.Message,
                    "שגיאה ב-SaveAsPDF",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
            finally
            {
                // Release COM objects to avoid memory leaks
                if (contactItems != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(contactItems);
                if (contactsFolder != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(contactsFolder);
                if (nameSpace != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(nameSpace);
                if (app != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }
            return output;
        }

        public void EnumerateFoldersInDefaultStore()
        {
            var root = Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder() as Folder;
            EnumerateFolders(root);
        }

        /// <summary>
        /// Uses recursion to enumerate Outlook sub-folders. 
        /// </summary>
        /// <param name="folder"></param>
        public void EnumerateFolders(Folder folder)
        {
            var childFolders = folder.Folders;
            int count = childFolders.Count;
            for (int i = 1; i <= count; i++)
            {
                var childFolder = childFolders[i] as Folder;
                if (childFolder != null)
                {
                    Debug.WriteLine(childFolder.FolderPath);
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
            var folderContacts = Globals.ThisAddIn.Application.ActiveExplorer().Session
                .GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            var searchFolder = folderContacts.Items;
            int counter = 0;
            int count = searchFolder.Count;
            for (int i = 1; i <= count; i++)
            {
                var foundContact = searchFolder[i] as ContactItem;
                if (foundContact != null && foundContact.LastName != null && foundContact.LastName.Contains(findLastName))
                {
                    foundContact.Display(false);
                    counter++;
                }
            }
            XMessageBox.Show(
                $"יש לך {counter} אנשי קשר עם שם משפחה המכיל את {findLastName}.",
                "תוצאות חיפוש אנשי קשר",
                XMessageBoxButtons.OK,
                XMessageBoxIcon.Information,
                XMessageAlignment.Right,
                XMessageLanguage.Hebrew
            );
        }

        public static void FindContact(string inString)
        {
            NameSpace outlookNameSpace = null;
            MAPIFolder contactsFolder = null;
            Items contactItems = null;
            try
            {
                outlookNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
                contactsFolder = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                contactItems = contactsFolder.Items;

                var contact = contactItems.Find(inString.Trim()) as ContactItem;
                if (contact != null)
                {
                    contact.Display(true);
                }
                else
                {
                    XMessageBox.Show(
                        "פרטי איש הקשר לא נמצאו.",
                        "שגיאה",
                        XMessageBoxButtons.OK,
                        XMessageBoxIcon.Warning,
                        XMessageAlignment.Right,
                        XMessageLanguage.Hebrew
                    );
                }
            }
            catch (Exception)
            {
                XMessageBox.Show(
                    "אירעה שגיאה במהלך חיפוש איש הקשר.",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
            finally
            {
                // Release COM objects to avoid memory leaks
                if (contactItems != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(contactItems);
                if (contactsFolder != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(contactsFolder);
                if (outlookNameSpace != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookNameSpace);
            }
        }
    }
}
