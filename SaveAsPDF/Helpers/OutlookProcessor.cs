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
            Table table = null;

            try
            {
                // Use the existing Outlook instance from the add-in
                nameSpace = AddinGlobals.OutlookApp.Session;
                contactsFolder = nameSpace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

                const string emailProp = "urn:schemas:contacts:email1";
                const string firstNameProp = "urn:schemas:contacts:givenName";
                const string lastNameProp = "urn:schemas:contacts:sn";

                // Use Table for significantly faster enumeration (no per-item COM object)
                table = contactsFolder.GetTable("[MessageClass] = 'IPM.Contact'");
                table.Columns.RemoveAll();
                table.Columns.Add(emailProp);
                table.Columns.Add(firstNameProp);
                table.Columns.Add(lastNameProp);

                while (!table.EndOfTable)
                {
                    Row row = table.GetNextRow();
                    string email = row[emailProp] as string;
                    if (!string.IsNullOrEmpty(email))
                    {
                        output.Add(new EmployeeModel
                        {
                            EmailAddress = email,
                            FirstName = row[firstNameProp] as string,
                            LastName = row[lastNameProp] as string
                        });
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(row);
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
                if (table != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
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
                root = AddinGlobals.OutlookApp.Session.DefaultStore.GetRootFolder() as Folder;
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
                session = AddinGlobals.OutlookApp.Session;
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

        /// <summary>
        /// Returns a tree of all Outlook contact folders across every store.
        /// Store-level nodes have a null EntryID; folder-level nodes carry their EntryID/StoreID.
        /// </summary>
        public static List<ContactFolderInfo> GetContactFolderTree()
        {
            var result = new List<ContactFolderInfo>();
            NameSpace nameSpace = null;
            Stores stores = null;
            try
            {
                nameSpace = AddinGlobals.OutlookApp.Session;
                stores = nameSpace.Stores;
                int storeCount = stores.Count;
                for (int s = 1; s <= storeCount; s++)
                {
                    Store store = null;
                    MAPIFolder rootFolder = null;
                    try
                    {
                        store = stores[s];
                        rootFolder = store.GetRootFolder();
                        var storeNode = new ContactFolderInfo
                        {
                            Name = store.DisplayName,
                            EntryID = null,
                            StoreID = store.StoreID
                        };
                        CollectContactFolders(rootFolder, storeNode);
                        if (storeNode.SubFolders.Count > 0)
                            result.Add(storeNode);
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Skip stores that cannot be accessed
                    }
                    finally
                    {
                        if (rootFolder != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rootFolder);
                        if (store != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(store);
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Could not enumerate stores
            }
            finally
            {
                if (stores != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(stores);
                if (nameSpace != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(nameSpace);
            }
            return result;
        }

        /// <summary>
        /// Recursively collects contact-type folders from the given parent folder.
        /// Hidden system folders and empty folders (no items and no visible sub-folders) are excluded.
        /// </summary>
        private static void CollectContactFolders(MAPIFolder parentFolder, ContactFolderInfo parentInfo)
        {
            Folders subFolders = null;
            try
            {
                subFolders = parentFolder.Folders;
                int count = subFolders.Count;
                for (int i = 1; i <= count; i++)
                {
                    MAPIFolder folder = null;
                    try
                    {
                        folder = subFolders[i];
                        if (folder.DefaultItemType == OlItemType.olContactItem)
                        {
                            // Skip hidden system folders
                            if (IsFolderHidden(folder))
                                continue;

                            var folderInfo = new ContactFolderInfo
                            {
                                Name = folder.Name,
                                EntryID = folder.EntryID,
                                StoreID = folder.StoreID
                            };
                            CollectContactFolders(folder, folderInfo);

                            // Skip empty folders that also have no visible sub-folders
                            if (IsFolderEmpty(folder) && folderInfo.SubFolders.Count == 0)
                                continue;

                            parentInfo.SubFolders.Add(folderInfo);
                        }
                        else
                        {
                            // Recurse into non-contact folders; contact folders may be nested
                            var temp = new ContactFolderInfo();
                            CollectContactFolders(folder, temp);
                            parentInfo.SubFolders.AddRange(temp.SubFolders);
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Skip inaccessible folders
                    }
                    finally
                    {
                        if (folder != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
                    }
                }
            }
            finally
            {
                if (subFolders != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(subFolders);
            }
        }

        /// <summary>
        /// Returns true if the folder has the MAPI PR_ATTR_HIDDEN property set.
        /// </summary>
        private static bool IsFolderHidden(MAPIFolder folder)
        {
            PropertyAccessor pa = null;
            try
            {
                const string PR_ATTR_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x10F4000B";
                pa = folder.PropertyAccessor;
                object val = pa.GetProperty(PR_ATTR_HIDDEN);
                return val is bool && (bool)val;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (pa != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(pa);
            }
        }

        /// <summary>
        /// Returns true if the folder contains no items.
        /// Uses folder.Items.Count which is a lightweight property access
        /// (avoids the expensive Restrict query used by GetContactCount).
        /// </summary>
        private static bool IsFolderEmpty(MAPIFolder folder)
        {
            Items items = null;
            try
            {
                items = folder.Items;
                return items.Count == 0;
            }
            catch
            {
                return true;
            }
            finally
            {
                if (items != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
            }
        }

        /// <summary>
        /// Returns the number of IPM.Contact items in the folder.
        /// </summary>
        private static int GetContactCount(MAPIFolder folder)
        {
            Items allItems = null;
            Items contactItems = null;
            try
            {
                allItems = folder.Items;
                contactItems = allItems.Restrict("[MessageClass] = 'IPM.Contact'");
                return contactItems.Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                if (contactItems != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(contactItems);
                if (allItems != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(allItems);
            }
        }

        /// <summary>
        /// Lists contacts from a specific Outlook folder identified by EntryID and StoreID.
        /// </summary>
        public static List<EmployeeModel> ListContactsFromFolder(string entryID, string storeID)
        {
            var output = new List<EmployeeModel>();
            NameSpace nameSpace = null;
            MAPIFolder folder = null;
            Table table = null;
            try
            {
                nameSpace = AddinGlobals.OutlookApp.Session;
                folder = nameSpace.GetFolderFromID(entryID, storeID);

                const string emailProp = "urn:schemas:contacts:email1";
                const string firstNameProp = "urn:schemas:contacts:givenName";
                const string lastNameProp = "urn:schemas:contacts:sn";

                // Use Table for significantly faster enumeration (no per-item COM object)
                table = folder.GetTable("[MessageClass] = 'IPM.Contact'");
                table.Columns.RemoveAll();
                table.Columns.Add(emailProp);
                table.Columns.Add(firstNameProp);
                table.Columns.Add(lastNameProp);

                while (!table.EndOfTable)
                {
                    Row row = table.GetNextRow();
                    string email = row[emailProp] as string;
                    if (!string.IsNullOrEmpty(email))
                    {
                        output.Add(new EmployeeModel
                        {
                            EmailAddress = email,
                            FirstName = row[firstNameProp] as string,
                            LastName = row[lastNameProp] as string
                        });
                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(row);
                }
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Folder could not be read
            }
            finally
            {
                if (table != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(table);
                if (folder != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
                if (nameSpace != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(nameSpace);
            }
            return output;
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
                outlookNameSpace = AddinGlobals.OutlookApp.Session;
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

