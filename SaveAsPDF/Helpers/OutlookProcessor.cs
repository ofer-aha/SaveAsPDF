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

            NameSpace nameSpace = null;
            MAPIFolder contactsFolder = null;
            Items contactItems = null;

            try
            {
                // Use the existing Outlook instance from the add-in
                nameSpace = Globals.ThisAddIn.Application.Session;
                contactsFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

                // Restrict to contacts only to avoid dist lists and other items
                contactItems = contactsFolder.Items.Restrict("[MessageClass] = 'IPM.Contact'");

                int count = contactItems.Count;
                for (int i = 1; i <= count; i++)
                {
                    var item = contactItems[i] as ContactItem;
                    try
                    {
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
                    finally
                    {
                        if (item != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
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
            }
            return output;
        }

        public void EnumerateFoldersInDefaultStore()
        {
            Folder root = null;
            try
            {
                root = Globals.ThisAddIn.Application.Session.DefaultStore.GetRootFolder() as Folder;
                EnumerateFolders(root);
            }
            finally
            {
                if (root != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(root);
            }
        }

        /// <summary>
        /// Uses recursion to enumerate Outlook sub-folders. 
        /// </summary>
        /// <param name="folder"></param>
        public void EnumerateFolders(Folder folder)
        {
            Folders childFolders = null;
            try
            {
                childFolders = folder.Folders;
                int count = childFolders.Count;
                for (int i = 1; i <= count; i++)
                {
                    var childFolder = childFolders[i] as Folder;
                    try
                    {
                        if (childFolder != null)
                        {
                            Debug.WriteLine(childFolder.FolderPath);
                            EnumerateFolders(childFolder);
                        }
                    }
                    finally
                    {
                        if (childFolder != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(childFolder);
                    }
                }
            }
            finally
            {
                if (childFolders != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(childFolders);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="findLastName"></param>
        public static void AccessContacts(string findLastName)
        {
            if (string.IsNullOrWhiteSpace(findLastName))
            {
                XMessageBox.Show(
                    "יש לספק טקסט לחיפוש.",
                    "חיפוש אנשי קשר",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Information,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return;
            }

            NameSpace session = null;
            MAPIFolder folderContacts = null;
            Items searchFolder = null;
            try
            {
                session = Globals.ThisAddIn.Application.Session;
                folderContacts = session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

                // Only contacts
                searchFolder = folderContacts.Items.Restrict("[MessageClass] = 'IPM.Contact'");

                int counter = 0;
                int count = searchFolder.Count;
                for (int i = 1; i <= count; i++)
                {
                    var foundContact = searchFolder[i] as ContactItem;
                    try
                    {
                        if (foundContact != null &&
                            !string.IsNullOrEmpty(foundContact.LastName) &&
                            foundContact.LastName.IndexOf(findLastName, System.StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            foundContact.Display(false);
                            counter++;
                        }
                    }
                    finally
                    {
                        if (foundContact != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(foundContact);
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
            finally
            {
                if (searchFolder != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(searchFolder);
                if (folderContacts != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(folderContacts);
                if (session != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(session);
            }
        }

        public static void FindContact(string inString)
        {
            if (string.IsNullOrWhiteSpace(inString))
            {
                XMessageBox.Show(
                    "יש לספק טקסט לחיפוש.",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return;
            }

            NameSpace outlookNameSpace = null;
            MAPIFolder contactsFolder = null;
            Items contactItems = null;
            Items results = null;
            ContactItem contact = null;
            try
            {
                outlookNameSpace = Globals.ThisAddIn.Application.Session;
                contactsFolder = outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

                contactItems = contactsFolder.Items.Restrict("[MessageClass] = 'IPM.Contact'");

                var term = inString.Trim();
                // Escape single quotes in filter
                var safe = term.Replace("'", "''");

                // Build a broad filter to match common fields
                string filter = $"[Email1Address] = '{safe}' OR [FullName] = '{safe}' OR [FirstName] = '{safe}' OR [LastName] = '{safe}'";

                results = contactItems.Restrict(filter);
                if (results != null && results.Count > 0)
                {
                    contact = results[1] as ContactItem;
                    if (contact != null)
                    {
                        contact.Display(true);
                        return;
                    }
                }

                // Fallback to manual scan with contains (case-insensitive)
                int count = contactItems.Count;
                for (int i = 1; i <= count; i++)
                {
                    var itm = contactItems[i] as ContactItem;
                    try
                    {
                        if (itm == null) continue;
                        if ((!string.IsNullOrEmpty(itm.Email1Address) && itm.Email1Address.Equals(term, System.StringComparison.OrdinalIgnoreCase)) ||
                            (!string.IsNullOrEmpty(itm.FullName) && itm.FullName.Equals(term, System.StringComparison.OrdinalIgnoreCase)) ||
                            (!string.IsNullOrEmpty(itm.LastName) && itm.LastName.IndexOf(term, System.StringComparison.OrdinalIgnoreCase) >= 0) ||
                            (!string.IsNullOrEmpty(itm.FirstName) && itm.FirstName.IndexOf(term, System.StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            itm.Display(true);
                            return;
                        }
                    }
                    finally
                    {
                        if (itm != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(itm);
                    }
                }

                XMessageBox.Show(
                    "פרטי איש הקשר לא נמצאו.",
                    "שגיאה",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Warning,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
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
                if (contact != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(contact);
                if (results != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(results);
                if (contactItems != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(contactItems);
                if (contactsFolder != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(contactsFolder);
                if (outlookNameSpace != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookNameSpace);
            }
        }
    }
}
