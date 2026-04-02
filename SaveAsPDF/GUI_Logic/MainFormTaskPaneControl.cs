using Microsoft.Office.Interop.Outlook;
using SaveAsPDF.Helpers;
using SaveAsPDF.Models;
using SaveAsPDF.Properties;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Exception = System.Exception;

namespace SaveAsPDF
{
    public partial class MainFormTaskPaneControl : UserControl, IEmployeeRequester, INewProjectRequester, ISettingsRequester
    {
        private ProjectModel _projectModel = new ProjectModel();
        public static SettingsModel settingsModel = new SettingsModel();

        private MailItem _mailItem;
        private string _currentProjectId;
        private bool _isDoubleClick;

        private readonly List<string> searchHistory = new List<string>();
        private readonly List<AttachmentsModel> attachmentsModels = new List<AttachmentsModel>();
        private readonly BindingList<EmployeeModel> _employeesBindingList = new BindingList<EmployeeModel>();
        private bool _selectingLeader;

        // Core controls roughly mirroring FormMain, but as a task pane
        private readonly TextBox txtProjectID = new TextBox();
        private readonly TextBox txtProjectName = new TextBox();
        private readonly TextBox txtProjectLeader = new TextBox();
        private readonly TextBox txtSubject = new TextBox();
        private readonly ExplorerAddressBar cmbSaveLocation = new ExplorerAddressBar();
        private readonly RichTextBox rtxtNotes = new RichTextBox();
        private readonly RichTextBox rtxtProjectNotes = new RichTextBox();
        private readonly DataGridView dgvAttachments = new DataGridView();
        private readonly DataGridView dgvEmployees = new DataGridView();
        private readonly TreeView tvFolders = new TreeView();
        private readonly CheckBox chkbSendNote = new CheckBox();
        private readonly CheckBox chkbSelectAllAttachments = new CheckBox();
        private readonly CheckBox chbOpenPDF = new CheckBox();
        private readonly Button btnOK = new Button();
        private readonly Button btnCancel = new Button();
        private readonly Button btnSettings = new Button();
        private readonly Button btnNewProject = new Button();
        private readonly Button btnFolders = new Button();
        private readonly Button btnCopyNotesToMail = new Button();
        private readonly Button btnCopyNotesToProject = new Button();
        private readonly Button btnStyle = new Button();
        private readonly Button btnRemoveEmployee = new Button();
        private readonly Button btnPhoneBook = new Button();
        private readonly Button btnProjectLeader = new Button();
        private TabControl tabNotes = new TabControl();
        private readonly TabControl tabFilesFolders = new TabControl();
        private readonly StatusStrip statusStrip = new StatusStrip();
        private readonly ToolStripStatusLabel tsslStatus = new ToolStripStatusLabel();
        private readonly ErrorProvider errorProviderMain = new ErrorProvider();
        private readonly FontDialog dlgFont = new FontDialog();

        public void LoadMailItem(MailItem mailItem)
        {
            if (ReferenceEquals(_mailItem, mailItem)) return;
            if (mailItem == null)
            {
                _mailItem = null;
                ClearMailRelatedUi();
                return;
            }
            _mailItem = mailItem;

            string htmlProjectId, htmlProjectName;
            bool loadedFromHtml = ExtractProjectFieldsFromHtmlBody(_mailItem, out htmlProjectId, out htmlProjectName);
            if (loadedFromHtml)
            {
                _currentProjectId = htmlProjectId;
                txtProjectID.Text = htmlProjectId;
                if (!string.IsNullOrWhiteSpace(htmlProjectName))
                    txtProjectName.Text = htmlProjectName;
                if (!string.IsNullOrWhiteSpace(htmlProjectId) && htmlProjectId.SafeProjectID())
                    ValidateAndLoadProjectById(htmlProjectId, showErrorDialogs: false);
                else
                    ClearProjectRelatedUi();

                // LoadProjectData inside ValidateAndLoadProjectById may overwrite
                // the project name with whatever is in the XML file. Restore the
                // HTML-embedded value so the user sees the correct name.
                if (!string.IsNullOrWhiteSpace(htmlProjectName))
                {
                    txtProjectName.Text = htmlProjectName;
                    if (_projectModel != null)
                        _projectModel.ProjectName = htmlProjectName;
                }
            }
            else
            {
                _currentProjectId = string.Empty;
                txtProjectID.Text = string.Empty;
                txtProjectName.Text = string.Empty;
                ClearProjectRelatedUi();
            }

            txtSubject.Text = LoadEmailSubject();
            ProcessMailItem(_mailItem);
        }

        private string ExtractProjectIdFromSubject(string subject)
        {
            if (string.IsNullOrWhiteSpace(subject)) return string.Empty;
            string combinedText = subject;
            try
            {
                if (_mailItem != null)
                {
                    string body = _mailItem.Body ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(body)) combinedText = subject + " " + body;
                }
            }
            catch { }
            var separators = new[] { ' ', '\t', '-', '_', ':', ';', ',', '[', ']', '(', ')', '{', '}' };
            var parts = combinedText.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
                if (part.SafeProjectID()) return part;
            return string.Empty;
        }

        private void ClearMailRelatedUi()
        {
            txtSubject.Clear();
            dgvAttachments.DataSource = null;
            attachmentsModels.Clear();
        }

        private void ClearProjectRelatedUi()
        {
            txtProjectName.Clear();
            cmbSaveLocation.Path = string.Empty;
            rtxtProjectNotes.Clear();
            tvFolders.Nodes.Clear();
        }

        private void ValidateAndLoadProjectById(string projectID, bool showErrorDialogs = true)
        {
            if (!projectID.SafeProjectID())
            {
                errorProviderMain.SetError(txtProjectID, "מספר פרויקט לא חוקי");
                tsslStatus.Text = errorProviderMain.GetError(txtProjectID);
                return;
            }
            errorProviderMain.SetError(txtProjectID, string.Empty);
            tsslStatus.Text = string.Empty;
            UpdateAutoCompleteSource(projectID);
            var projectRootFolder = projectID.ProjectFullPath(settingsModel.RootDrive);
            // Refresh to get current state from the file system (important for network drives)
            projectRootFolder.Refresh();
            if (!projectRootFolder.Exists)
            {
                if (showErrorDialogs)
                {
                    var createResult = MessageBox.Show(
                        "הפרויקט לא קיים. האם ליצור תיקיית פרויקט חדשה?",
                        "SaveAsPDF",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question,
                        MessageBoxDefaultButton.Button2,
                        MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);

                    if (createResult == DialogResult.Yes)
                    {
                        try
                        {
                            Directory.CreateDirectory(projectRootFolder.FullName);
                        }
                        catch (Exception ex)
                        {
                            // On network drives CreateDirectory may throw even when
                            // the folder already exists. Re-check before giving up.
                            projectRootFolder.Refresh();
                            if (!projectRootFolder.Exists)
                            {
                                XMessageBox.Show($"שגיאה ביצירת תיקיית פרויקט: {ex.Message}", "שגיאה", XMessageBoxButtons.OK, XMessageBoxIcon.Error, XMessageAlignment.Right, XMessageLanguage.Hebrew);
                                ClearProjectRelatedUi();
                                return;
                            }
                        }
                    }
                    else
                    {
                        ClearProjectRelatedUi();
                        return;
                    }
                }
                else
                {
                    ClearProjectRelatedUi();
                    return;
                }
            }
            settingsModel = SettingsHelpers.LoadProjectSettings(projectID);
            LoadProjectData();
            LoadEmployeeData();
            UpdateUI();
        }

        private void LoadProjectData()
        {
            if (File.Exists(settingsModel.XmlProjectFile))
            {
                try
                {
                    _projectModel = XmlFileHelper.XmlProjectFileToModel(settingsModel.XmlProjectFile);
                    if (_projectModel != null)
                    {
                        txtProjectName.Text = _projectModel.ProjectName;
                        rtxtProjectNotes.Text = _projectModel.ProjectNotes;
                    }
                }
                catch (Exception ex)
                {
                    XMessageBox.Show($"שגיאה בטעינת נתוני הפרויקט: {ex.Message}", "SaveAsPDF:LoadProjectData", XMessageBoxButtons.OK, XMessageBoxIcon.Warning, XMessageAlignment.Right, XMessageLanguage.Hebrew);
                    SetDefaultProjectModel();
                }
            }
            else SetDefaultProjectModel();
        }

        private void SetDefaultProjectModel()
        {
            _projectModel = new ProjectModel
            {
                ProjectName = "פרויקט ברירת מחדל",
                ProjectNumber = txtProjectID.Text,
                NoteToProjectLeader = false,
                DefaultSaveFolder = settingsModel.DefaultSavePath,
                ProjectNotes = "הערות ברירת מחדל",
                LastSavePath = settingsModel.DefaultSavePath
            };
            if (!string.IsNullOrEmpty(settingsModel.XmlProjectFile))
                settingsModel.XmlProjectFile.ProjectModelToXmlFile(_projectModel);
            txtProjectName.Text = _projectModel.ProjectName;
            rtxtProjectNotes.Text = _projectModel.ProjectNotes;
        }

        private void LoadEmployeeData()
        {
            _employeesBindingList.Clear();
            if (File.Exists(settingsModel.XmlEmployeesFile))
            {
                var loaded = settingsModel.XmlEmployeesFile.XmlEmployeesFileToModel();
                if (loaded != null)
                    foreach (var em in loaded) _employeesBindingList.Add(em);
                var leader = _employeesBindingList.FirstOrDefault(e => e.IsLeader);
                if (leader != null)
                {
                    var first = leader.FirstName ?? string.Empty;
                    var last = leader.LastName ?? string.Empty;
                    string fullName = ($"{first} {last}").Trim();
                    if (string.IsNullOrWhiteSpace(fullName)) fullName = leader.EmailAddress ?? string.Empty;
                    txtProjectLeader.Text = fullName;
                }
                else txtProjectLeader.Clear();
            }
            else
            {
                if (!string.IsNullOrEmpty(settingsModel.XmlEmployeesFile))
                    settingsModel.XmlEmployeesFile.EmployeesModelToXmlFile(new List<EmployeeModel>());
                txtProjectLeader.Clear();
            }
        }

        private void UpdateUI()
        {
            try
            {
                string targetPath = settingsModel.DefaultSavePath;
                if (string.IsNullOrEmpty(targetPath) && settingsModel.ProjectRootFolder != null && settingsModel.ProjectRootFolder.Exists)
                    targetPath = settingsModel.ProjectRootFolder.FullName;
                if (!string.IsNullOrEmpty(targetPath))
                {
                    cmbSaveLocation.Path = targetPath;
                }
                tvFolders.Nodes.Clear();
                if (settingsModel.ProjectRootFolder != null && settingsModel.ProjectRootFolder.Exists)
                {
                    var rootDir = settingsModel.ProjectRootFolder;
                    var rootNode = new TreeNode(rootDir.Name) { Tag = rootDir.FullName };
                    rootNode.Nodes.Add("...");
                    tvFolders.Nodes.Add(rootNode);
                    tvFolders.SelectedNode = tvFolders.Nodes[0];
                }
                else
                {
                    XMessageBox.Show("תיקיית השורש של הפרויקט אינה קיימת.", "שגיאה", XMessageBoxButtons.OK, XMessageBoxIcon.Warning, XMessageAlignment.Right, XMessageLanguage.Hebrew);
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show($"אירעה שגיאה בעת עדכון הממשק: {ex.Message}", "שגיאה", XMessageBoxButtons.OK, XMessageBoxIcon.Error, XMessageAlignment.Right, XMessageLanguage.Hebrew);
            }
        }

        private void ApplySystemColors()
        {
            BackColor = SystemColors.Control;
            ForeColor = SystemColors.ControlText;
        }

        private string LoadEmailSubject() => _mailItem?.Subject ?? "Default Email Subject";

        private void SetupContextMenus()
        {
            rtxtNotes.EnableContextMenu();
            rtxtProjectNotes.EnableContextMenu();
            txtProjectID.EnableContextMenu();
            txtProjectName.EnableContextMenu();
            tvFolders.EnableContextMenu(() => settingsModel.ProjectRootFolder);
        }

        private void ConfigureProjectIdAutoComplete()
        {
            txtProjectID.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtProjectID.AutoCompleteSource = AutoCompleteSource.CustomSource;
        }

        private void PopulateDriveNodes()
        {
            tvFolders.Nodes.Clear();
            foreach (string drive in Environment.GetLogicalDrives())
            {
                try
                {
                    var di = new DriveInfo(drive);
                    var node = new TreeNode(drive.Substring(0, 1)) { Tag = drive };
                    if (di.IsReady) node.Nodes.Add("...");
                    tvFolders.Nodes.Add(node);
                }
                catch { }
            }
        }

        private void UpdateAutoCompleteSource(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return;
            searchHistory.Remove(text);
            searchHistory.Insert(0, text);
            int maxCount = 10;
            if (searchHistory.Count > maxCount) searchHistory.RemoveRange(maxCount, searchHistory.Count - maxCount);
            if (Settings.Default.LastProjects == null) Settings.Default.LastProjects = new StringCollection();
            Settings.Default.LastProjects.Clear();
            Settings.Default.LastProjects.AddRange(searchHistory.ToArray());
            Settings.Default.Save();
            txtProjectID.AutoCompleteCustomSource.Clear();
            txtProjectID.AutoCompleteCustomSource.AddRange(searchHistory.ToArray());
        }

        private void LoadSearchHistory()
        {
            if (Settings.Default.LastProjects == null) Settings.Default.LastProjects = new StringCollection();
            searchHistory.Clear();
            searchHistory.AddRange(Settings.Default.LastProjects.Cast<string>().Distinct());
            txtProjectID.AutoCompleteCustomSource.Clear();
            txtProjectID.AutoCompleteCustomSource.AddRange(searchHistory.ToArray());
        }

        private void btnCancel_Click(object sender, EventArgs e) { }

        private void btnFolders_Click(object sender, EventArgs e)
        {
            var dialog = new FolderPicker
            {
                InputPath = settingsModel.ProjectRootFolder != null && settingsModel.ProjectRootFolder.Exists ? settingsModel.ProjectRootFolder.FullName : settingsModel.RootDrive
            };
            if (dialog.ShowDialog(Handle) == true)
            {
                tvFolders.Nodes.Clear();
                var root = new TreeNode(new DirectoryInfo(dialog.ResultPath).Name) { Tag = dialog.ResultPath };
                root.Nodes.Add("...");
                tvFolders.Nodes.Add(root);
                cmbSaveLocation.Path = dialog.ResultPath;
            }
        }

        private void txtProjectID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) btnOK.Select();
        }

        private void txtProjectID_Validating(object sender, CancelEventArgs e)
        {
            if (!txtProjectID.Text.SafeProjectID())
            {
                errorProviderMain.SetError(txtProjectID, "מספר פרויקט לא חוקי");
                txtProjectID.Select(0, txtProjectID.Text.Length);
                tsslStatus.Text = errorProviderMain.GetError(txtProjectID);
                e.Cancel = true;
            }
        }

        private void txtProjectID_Validated(object sender, EventArgs e)
        {
            errorProviderMain.SetError(txtProjectID, string.Empty);
            tsslStatus.Text = errorProviderMain.GetError(txtProjectID);
            string projectID = txtProjectID.Text;
            if (!string.IsNullOrWhiteSpace(projectID))
            {
                ValidateAndLoadProjectById(projectID);
                if (_mailItem != null) ProcessMailItem(_mailItem);
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            string sPath = cmbSaveLocation.Path;
            if (string.IsNullOrEmpty(sPath))
            {
                var dialog = new FolderPicker { InputPath = settingsModel.RootDrive };
                if (dialog.ShowDialog(Handle) == true)
                {
                    sPath = dialog.ResultPath;
                    cmbSaveLocation.Path = sPath;
                }
            }
            if (!string.IsNullOrEmpty(sPath))
            {
                try
                {
                    sPath = Path.GetFullPath(sPath);
                    if (!Path.IsPathRooted(sPath))
                        throw new ArgumentException("Path is not rooted");
                }
                catch (ArgumentException)
                {
                    XMessageBox.Show("הנתיב מכיל תווים לא חוקיים או אינו נתיב מלא. אנא בחר נתיב תקין.", "שגיאה", XMessageBoxButtons.OK, XMessageBoxIcon.Error, XMessageAlignment.Right, XMessageLanguage.Hebrew);
                    return;
                }
                catch (Exception ex)
                {
                    XMessageBox.Show($"נתיב לא תקין: {ex.Message}", "שגיאה", XMessageBoxButtons.OK, XMessageBoxIcon.Error, XMessageAlignment.Right, XMessageLanguage.Hebrew);
                    return;
                }
                var directory = new DirectoryInfo(sPath);
                if (!directory.Exists) FileFoldersHelper.CreateDirectory(directory.FullName);
            }
            else
            {
                XMessageBox.Show("יש לבחור או לציין מיקום שמירה תקין.", "שגיאה", XMessageBoxButtons.OK, XMessageBoxIcon.Error, XMessageAlignment.Right, XMessageLanguage.Hebrew);
                return;
            }

            List<AttachmentsModel> savedAttachmentsModels = new List<AttachmentsModel>();
            try
            {
                var saved = _mailItem.SaveAttachments(dgvAttachments, sPath, overWrite: false);
                int idx = 0;
                foreach (var entry in saved)
                {
                    var parts = entry.Split(new[] { '|' }, 2);
                    var fname = parts.Length > 0 ? parts[0] : string.Empty;
                    var fsize = parts.Length > 1 ? parts[1] : string.Empty;
                    savedAttachmentsModels.Add(new AttachmentsModel { attachmentId = idx++, isChecked = true, fileName = fname, fileSize = fsize });
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show($"שגיאה בשמירת קבצים מצורפים: {ex.Message}", "SaveAsPDF", XMessageBoxButtons.OK, XMessageBoxIcon.Warning, XMessageAlignment.Right, XMessageLanguage.Hebrew);
            }

            string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string sanitizedProjectName = txtProjectName.Text.SafeFolderName();
            string htmlFileName = $"{sanitizedProjectName}_{timeStamp}.html";
            string htmlFilePath = Path.Combine(sPath, htmlFileName);

            // Pre-compute the PDF filename so the HTML link and the actual PDF file match
            string mailSubjectText = (_mailItem != null && !string.IsNullOrWhiteSpace(_mailItem.Subject)) ? _mailItem.Subject : txtSubject.Text;
            string pdfFileName = $"{timeStamp}_{mailSubjectText.SafeFolderName()}.pdf";

            try
            {
                HtmlHelper.GenerateHtmlToFile(htmlFilePath, sPath, _employeesBindingList.ToList(), (savedAttachmentsModels.Count > 0) ? savedAttachmentsModels : attachmentsModels, txtProjectName.Text, txtProjectID.Text, rtxtNotes.Text, Environment.UserName, mailSubjectText, pdfFileName);
                try { string htmlContent = File.ReadAllText(htmlFilePath); _mailItem.HTMLBody = htmlContent + _mailItem.HTMLBody; }
                catch (Exception ex)
                {
                    XMessageBox.Show($"שגיאה בטעינת קובץ HTML שנוצר: {ex.Message}", "SaveAsPDF", XMessageBoxButtons.OK, XMessageBoxIcon.Warning, XMessageAlignment.Right, XMessageLanguage.Hebrew);
                }
                finally { try { if (File.Exists(htmlFilePath)) File.Delete(htmlFilePath); } catch { } }
            }
            catch (Exception ex)
            {
                XMessageBox.Show($"שגיאה ביצירת קובץ ה-HTML: {ex.Message}", "SaveAsPDF", XMessageBoxButtons.OK, XMessageBoxIcon.Error, XMessageAlignment.Right, XMessageLanguage.Hebrew);
                return;
            }
            // First: all HTMLBody modifications are done
            // (the htmlContent prepend already happened above)

            // Embed project fields BEFORE SaveToPDF
            EmbedProjectFieldsInHtmlBody(_mailItem, txtProjectID.Text, txtProjectName.Text);

            // Save to commit all HTML changes before PDF export
            _mailItem.Save();

            // Now export to PDF (on a cleanly saved message)
            _mailItem.SaveToPDF(sPath, pdfFileName);

            if (chkbSendNote.Checked)
            {
                MailItem forwardedItem = null;
                try
                {
                    var leader = _employeesBindingList.FirstOrDefault(emp => emp.IsLeader && !string.IsNullOrWhiteSpace(emp.EmailAddress));
                    var leaderEmail = leader?.EmailAddress;
                    if (!string.IsNullOrWhiteSpace(leaderEmail))
                    {
                        forwardedItem = _mailItem.Forward();
                        forwardedItem.To = leaderEmail;
                        forwardedItem.Send();
                    }
                    else XMessageBox.Show("לא נמצא אימייל של ראש הפרויקט. יש לבחור ראש פרויקט או לעדכן את כתובת האימייל שלו.", "SaveAsPDF", XMessageBoxButtons.OK, XMessageBoxIcon.Warning, XMessageAlignment.Right, XMessageLanguage.Hebrew);
                }
                catch (Exception ex)
                {
                    XMessageBox.Show($"שגיאה בשליחת הודעת העברה לראש הפרויקט: {ex.Message}", "SaveAsPDF", XMessageBoxButtons.OK, XMessageBoxIcon.Error, XMessageAlignment.Right, XMessageLanguage.Hebrew);
                }
                finally
                {
                    if (forwardedItem != null)
                    {
                        Marshal.ReleaseComObject(forwardedItem);
                        forwardedItem = null;
                    }
                }
            }
            if (_mailItem != null)
            {
                Marshal.ReleaseComObject(_mailItem);
                _mailItem = null;
                ClearMailRelatedUi();
            }
            if (chbOpenPDF.Checked)
            {
                string pdfFilePath = Path.Combine(sPath, pdfFileName);
                if (File.Exists(pdfFilePath)) System.Diagnostics.Process.Start(pdfFilePath);
                else XMessageBox.Show("קובץ ה-PDF לא נמצא.", "שגיאה", XMessageBoxButtons.OK, XMessageBoxIcon.Error, XMessageAlignment.Right, XMessageLanguage.Hebrew);
            }
        }

        private void btnRemoveEmployee_Click(object sender, EventArgs e)
        {
            var row = dgvEmployees.CurrentRow;
            if (row == null)
                return;

            var employee = row.DataBoundItem as EmployeeModel;
            if (employee == null)
                return;

            _employeesBindingList.Remove(employee);

            // If the removed employee was the leader, clear the leader text
            if (employee.IsLeader)
                txtProjectLeader.Clear();

            // Persist the updated employee list to the XML file
            if (!string.IsNullOrEmpty(settingsModel.XmlEmployeesFile))
                settingsModel.XmlEmployeesFile.EmployeesModelToXmlFile(_employeesBindingList.ToList());
        }

        // No-op: project fields are now embedded as visible HTML elements
        // (with id attributes) by HtmlHelper, so no hidden tags are needed.
        // Outlook 365 strips HTML comments and <meta> tags when sending,
        // but preserves id attributes on visible elements.
        private void EmbedProjectFieldsInHtmlBody(MailItem mailItem, string projectId, string projectName) { }

        private bool ExtractProjectFieldsFromHtmlBody(MailItem mailItem, out string projectId, out string projectName)
        {
            projectId = string.Empty;
            projectName = string.Empty;
            if (mailItem == null) return false;
            try
            {
                string htmlBody = mailItem.HTMLBody ?? string.Empty;

                // 1. Try id-attributed visible elements (current format, survives Outlook 365 send)
                //    <span id="SaveAsPDF-ProjectID">1000</span>
                var idElem = Regex.Match(htmlBody,
                    @"<span\s[^>]*id=""SaveAsPDF-ProjectID""[^>]*>([^<]*)</span>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (idElem.Success)
                    projectId = System.Net.WebUtility.HtmlDecode(idElem.Groups[1].Value).Trim();

                //    <h1 id="SaveAsPDF-ProjectName">Project Name</h1>
                var nameElem = Regex.Match(htmlBody,
                    @"<h1\s[^>]*id=""SaveAsPDF-ProjectName""[^>]*>([^<]*)</h1>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (nameElem.Success)
                    projectName = System.Net.WebUtility.HtmlDecode(nameElem.Groups[1].Value).Trim();

                // 2. Fallback: parse visible Hebrew-labelled text from the generated HTML
                //    "מספר פרויקט: 1000" in any element
                if (string.IsNullOrWhiteSpace(projectId))
                {
                    var idText = Regex.Match(htmlBody, @"מספר פרויקט:\s*(?:<[^>]+>)*\s*([^<]+)");
                    if (idText.Success)
                    {
                        var candidate = idText.Groups[1].Value.Trim();
                        if (candidate.SafeProjectID())
                            projectId = System.Net.WebUtility.HtmlDecode(candidate);
                    }
                }

                // 3. Legacy fallback: <meta> tags (may survive in local drafts)
                if (string.IsNullOrWhiteSpace(projectId))
                {
                    var idMeta = Regex.Match(htmlBody, @"<meta\s+name=""SaveAsPDF-ProjectID""\s+content=""([^""]*)""\s*/?>", RegexOptions.IgnoreCase);
                    if (idMeta.Success)
                        projectId = System.Net.WebUtility.HtmlDecode(idMeta.Groups[1].Value).Trim();
                }
                if (string.IsNullOrWhiteSpace(projectName))
                {
                    var nameMeta = Regex.Match(htmlBody, @"<meta\s+name=""SaveAsPDF-ProjectName""\s+content=""([^""]*)""\s*/?>", RegexOptions.IgnoreCase);
                    if (nameMeta.Success)
                        projectName = System.Net.WebUtility.HtmlDecode(nameMeta.Groups[1].Value).Trim();
                }

                // 4. Legacy fallback: HTML comments (may survive in local drafts)
                if (string.IsNullOrWhiteSpace(projectId))
                {
                    var idComment = Regex.Match(htmlBody, @"<!--\s*SaveAsPDF:ProjectID=(.*?)\s*-->");
                    if (idComment.Success)
                        projectId = idComment.Groups[1].Value.Trim();
                }
                if (string.IsNullOrWhiteSpace(projectName))
                {
                    var nameComment = Regex.Match(htmlBody, @"<!--\s*SaveAsPDF:ProjectName=(.*?)\s*-->");
                    if (nameComment.Success)
                        projectName = nameComment.Groups[1].Value.Trim();
                }

                return !string.IsNullOrWhiteSpace(projectId);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to extract project fields from HTML: {ex.Message}");
                return false;
            }
        }

        public void SettingsComplete(SettingsModel settings) => settingsModel = SettingsHelpers.LoadSettingsToModel(settings);

        public void EmployeeComplete(EmployeeModel model)
        {
            if (model.IsLeader)
            {
                foreach (var emp in _employeesBindingList)
                    emp.IsLeader = false;

                var existing = _employeesBindingList.FirstOrDefault(e => e.EmailAddress == model.EmailAddress);
                if (existing != null)
                    existing.IsLeader = true;
                else
                    _employeesBindingList.Add(model);

                var fullName = ($"{model.FirstName} {model.LastName}").Trim();
                if (string.IsNullOrWhiteSpace(fullName)) fullName = model.EmailAddress ?? string.Empty;
                txtProjectLeader.Text = fullName;
            }
            else
            {
                if (!_employeesBindingList.Any(e => e.EmailAddress == model.EmailAddress))
                    _employeesBindingList.Add(model);
            }

            // Persist the updated employee list to the XML file
            if (!string.IsNullOrEmpty(settingsModel.XmlEmployeesFile))
                settingsModel.XmlEmployeesFile.EmployeesModelToXmlFile(_employeesBindingList.ToList());
        }

        public void NewProjectComplete(ProjectModel model)
        {
            _projectModel = model;
            txtProjectID.Text = _projectModel.ProjectNumber;
            txtProjectName.Text = _projectModel.ProjectName;
            rtxtProjectNotes.Text = _projectModel.ProjectNotes;
        }

        private void btnProjectLeader_Click(object sender, EventArgs e)
        {
            _selectingLeader = true;
            using (var frmContacts = new FormContacts(this, true)) { frmContacts.ShowDialog(); }
            _selectingLeader = false;
        }

        private void ProcessMailItem(MailItem mailItem)
        {
            if (mailItem == null) return;
            txtSubject.Text = mailItem.Subject;
            var attachments = mailItem.GetAttachmentsFromEmail();
            int i = 0; attachmentsModels.Clear();
            foreach (var attachment in attachments)
            {
                if (attachment != null)
                {
                    attachmentsModels.Add(new AttachmentsModel { attachmentId = i++, isChecked = true, fileName = attachment.FileName, fileSize = attachment.Size.BytesToString() });
                }
            }
            dgvAttachments.DataSource = null;
            dgvAttachments.DataSource = attachmentsModels;
            if (dgvAttachments.Columns.Count > 0)
            {
                dgvAttachments.Columns[0].Visible = false;
                if (dgvAttachments.Columns.Count > 1) { dgvAttachments.Columns[1].HeaderText = "V"; dgvAttachments.Columns[1].ReadOnly = false; }
                if (dgvAttachments.Columns.Count > 2) { dgvAttachments.Columns[2].HeaderText = "שם קובץ"; dgvAttachments.Columns[2].ReadOnly = true; }
                if (dgvAttachments.Columns.Count > 3) { dgvAttachments.Columns[3].HeaderText = "גודל"; dgvAttachments.Columns[3].ReadOnly = true; }
            }
        }

        private void MouseEnterStatus(object sender, EventArgs e)
        {
            var ctl = sender as Control;
            if (ctl != null) tsslStatus.Text = ctl.Tag as string ?? string.Empty;
        }

        private void MouseLeaveStatus(object sender, EventArgs e) => tsslStatus.Text = string.Empty;

        private void WireStatusHelp()
        {
            btnOK.Tag = "שמור ל-PDF"; btnOK.MouseEnter += MouseEnterStatus; btnOK.MouseLeave += MouseLeaveStatus;
            btnSettings.Tag = "הגדרות"; btnSettings.MouseEnter += MouseEnterStatus; btnSettings.MouseLeave += MouseLeaveStatus;
            btnNewProject.Tag = "פרויקט חדש"; btnNewProject.MouseEnter += MouseEnterStatus; btnNewProject.MouseLeave += MouseLeaveStatus;
            btnFolders.Tag = "בחר תיקיית שורש"; btnFolders.MouseEnter += MouseEnterStatus; btnFolders.MouseLeave += MouseLeaveStatus;
            btnCopyNotesToMail.Tag = "העתק הערות לפרויקט אל המייל"; btnCopyNotesToMail.MouseEnter += MouseEnterStatus; btnCopyNotesToMail.MouseLeave += MouseLeaveStatus;
            btnCopyNotesToProject.Tag = "העתק ערות מהמייל אל הפרויקט"; btnCopyNotesToProject.MouseEnter += MouseEnterStatus; btnCopyNotesToProject.MouseLeave += MouseLeaveStatus;
            btnStyle.Tag = "בחר גופן להערות"; btnStyle.MouseEnter += MouseEnterStatus; btnStyle.MouseLeave += MouseLeaveStatus;
            btnRemoveEmployee.Tag = "הסר עובד מהרשימה"; btnRemoveEmployee.MouseEnter += MouseEnterStatus; btnRemoveEmployee.MouseLeave += MouseLeaveStatus;
            btnPhoneBook.Tag = "בחר עובד מספר טלפונים"; btnPhoneBook.MouseEnter += MouseEnterStatus; btnPhoneBook.MouseLeave += MouseLeaveStatus;
            btnProjectLeader.Tag = "בחר מתכנן מוביל בפרויקט"; btnProjectLeader.MouseEnter += MouseEnterStatus; btnProjectLeader.MouseLeave += MouseLeaveStatus;
            txtProjectLeader.Tag = "מתכנן מוביל בפרויקט"; txtProjectLeader.MouseEnter += MouseEnterStatus; txtProjectLeader.MouseLeave += MouseLeaveStatus;
            txtProjectID.Tag = "הכנס מספר פרויקט"; txtProjectID.MouseEnter += MouseEnterStatus; txtProjectID.MouseLeave += MouseLeaveStatus;
            txtProjectName.Tag = "שם הפרויקט"; txtProjectName.MouseEnter += MouseEnterStatus; txtProjectName.MouseLeave += MouseLeaveStatus;
            txtSubject.Tag = "נושא ההודעה"; txtSubject.MouseEnter += MouseEnterStatus; txtSubject.MouseLeave += MouseLeaveStatus;
            cmbSaveLocation.Tag = "בחר מיקום שמירה"; cmbSaveLocation.MouseEnter += MouseEnterStatus; cmbSaveLocation.MouseLeave += MouseLeaveStatus;
            rtxtNotes.Tag = "הערות למייל"; rtxtNotes.MouseEnter += MouseEnterStatus; rtxtNotes.MouseLeave += MouseLeaveStatus;
            rtxtProjectNotes.Tag = "הערות בפרויקט"; rtxtProjectNotes.MouseEnter += MouseEnterStatus; rtxtProjectNotes.MouseLeave += MouseLeaveStatus;
            chkbSendNote.Tag = "שלח ההערה לראש הפרויקט"; chkbSendNote.MouseEnter += MouseEnterStatus; chkbSendNote.MouseLeave += MouseLeaveStatus;
            chkbSelectAllAttachments.Tag = "בחר/הסר כל הקבצים"; chkbSelectAllAttachments.MouseEnter += MouseEnterStatus; chkbSelectAllAttachments.MouseLeave += MouseLeaveStatus;
            chbOpenPDF.Tag = "פתח PDF לאחר שמירה"; chbOpenPDF.MouseEnter += MouseEnterStatus; chbOpenPDF.MouseLeave += MouseLeaveStatus;
            tvFolders.Tag = "עץ תיקיות פרויקט"; tvFolders.MouseEnter += MouseEnterStatus; tvFolders.MouseLeave += MouseLeaveStatus;
            dgvAttachments.Tag = "קבצים מצורפים"; dgvAttachments.MouseEnter += MouseEnterStatus; dgvAttachments.MouseLeave += MouseLeaveStatus;
            dgvEmployees.Tag = "עובדים בפרויקט"; dgvEmployees.MouseEnter += MouseEnterStatus; dgvEmployees.MouseLeave += MouseLeaveStatus;
            tabNotes.Tag = "כרטיסיות הערות"; tabNotes.MouseEnter += MouseEnterStatus; tabNotes.MouseLeave += MouseLeaveStatus;
            tabFilesFolders.Tag = "קבצים ותיקיות"; tabFilesFolders.MouseEnter += MouseEnterStatus; tabFilesFolders.MouseLeave += MouseLeaveStatus;
        }

        private void dgvAttachments_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvAttachments.CurrentCell != null)
            {
                XMessageBox.Show(dgvAttachments.CurrentCell.Value?.ToString() ?? string.Empty, "פרטי קובץ", XMessageBoxButtons.OK, XMessageBoxIcon.Information, XMessageAlignment.Right, XMessageLanguage.Hebrew);
            }
        }

        private void btnStyle_Click(object sender, EventArgs e)
        {
            if (dlgFont.ShowDialog() == DialogResult.OK) rtxtNotes.SelectionFont = dlgFont.Font;
        }

        private void btnPhoneBook_Click(object sender, EventArgs e)
        {
            using (var frmContacts = new FormContacts(this)) { frmContacts.ShowDialog(); }
        }

        private void chkbSelectAllAttachments_CheckedChanged(object sender, EventArgs e)
        {
            chkbSelectAllAttachments.Text = chkbSelectAllAttachments.Checked ? "הסר הכל" : "בחר הכל";
            if (dgvAttachments.RowCount != 0)
            {
                dgvAttachments.SuspendLayout();
                try
                {
                    dgvAttachments.BeginEdit(false);
                    foreach (DataGridViewRow row in dgvAttachments.Rows) row.Cells[1].Value = chkbSelectAllAttachments.Checked;
                    dgvAttachments.EndEdit();
                }
                finally { dgvAttachments.ResumeLayout(false); }
            }
        }

        private void btnCopyNotesToMail_Click(object sender, EventArgs e) => rtxtNotes.Text += "\n " + rtxtProjectNotes.Text;

        private void btnCopyNotesToProject_Click(object sender, EventArgs e) => rtxtProjectNotes.Text += "\n " + rtxtNotes.Text;

        private void tvFolders_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            if (_isDoubleClick && e.Action == TreeViewAction.Expand) { e.Cancel = true; return; }
            if (e.Node.Nodes.Count == 1 && e.Node.Nodes[0].Text == "...")
            {
                try
                {
                    string basePath = e.Node.Tag as string;
                    if (!string.IsNullOrEmpty(basePath) && Directory.Exists(basePath))
                    {
                        e.Node.Nodes.Clear();
                        var nodes = TreeHelpers.GetFolderNodes(basePath, expanded: false);
                        foreach (var n in nodes) e.Node.Nodes.Add(n);
                    }
                }
                catch (Exception ex)
                {
                    XMessageBox.Show($"שגיאה בטעינת תיקיות: {ex.Message}", "SaveAsPDF:tvFolders_BeforeExpand", XMessageBoxButtons.OK, XMessageBoxIcon.Error, XMessageAlignment.Right, XMessageLanguage.Hebrew);
                }
            }
        }

        private void tvFolders_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            if (_isDoubleClick && e.Action == TreeViewAction.Collapse) e.Cancel = true;
        }

        private void tvFolders_MouseDown(object sender, MouseEventArgs e) => _isDoubleClick = e.Clicks > 1;

        private void tvFolders_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Tag is string basePath)
            {
                cmbSaveLocation.Path = basePath;
            }
        }

        private void tvFolders_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Tag is string basePath)
            {
                System.Diagnostics.Process.Start("explorer.exe", basePath);
                cmbSaveLocation.Path = basePath;
            }
        }

        private void dgvEmployees_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvEmployees.IsCurrentCellDirty && dgvEmployees.CurrentCell is DataGridViewCheckBoxCell) dgvEmployees.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void dgvEmployees_CellValueChanged(object sender, DataGridViewCellEventArgs e) { }

        private void MainFormTaskPaneControl_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) txtProjectID.Clear();
        }

        private void ConfigureEmployeeDataGrid()
        {
            dgvEmployees.AutoGenerateColumns = false;
            dgvEmployees.Columns.Clear();
            dgvEmployees.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn { Name = "Id", DataPropertyName = "Id", Visible = false },
                new DataGridViewTextBoxColumn { Name = "FirstName", DataPropertyName = "FirstName", HeaderText = "שם פרטי" },
                new DataGridViewTextBoxColumn { Name = "LastName", DataPropertyName = "LastName", HeaderText = "שם משפחה" },
                new DataGridViewTextBoxColumn { Name = "EmailAddress", DataPropertyName = "EmailAddress", HeaderText = "אימייל" }
            });
            dgvEmployees.DataBindingComplete -= DgvEmployees_DataBindingComplete;
            dgvEmployees.DataBindingComplete += DgvEmployees_DataBindingComplete;
            dgvEmployees.DataSource = _employeesBindingList;
            HideEmployeeIdColumn();
            dgvEmployees.ReadOnly = true;
            dgvEmployees.BackgroundColor = SystemColors.Window;
            dgvEmployees.ForeColor = SystemColors.WindowText;
            dgvEmployees.DefaultCellStyle.BackColor = SystemColors.Window;
            dgvEmployees.DefaultCellStyle.ForeColor = SystemColors.WindowText;
            dgvEmployees.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
            dgvEmployees.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;
            dgvEmployees.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.Control;
            dgvEmployees.AlternatingRowsDefaultCellStyle.ForeColor = SystemColors.ControlText;
            dgvEmployees.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dgvEmployees.ColumnHeadersDefaultCellStyle.ForeColor = SystemColors.ControlText;
            dgvEmployees.EnableHeadersVisualStyles = false;

            // Context menu for employee management
            var ctxEmployees = new ContextMenuStrip { RightToLeft = RightToLeft.Yes };
            var mnuAddEmployee = new ToolStripMenuItem("הוסף עובד");        // "Add Employee"
            var mnuRemoveEmployee = new ToolStripMenuItem("הסר עובד");       // "Remove Employee"
            var mnuSetLeader = new ToolStripMenuItem("קבע כמתכנן מוביל");    // "Set as Project Leader"
            mnuAddEmployee.Click += CtxEmployees_AddEmployee_Click;
            mnuRemoveEmployee.Click += CtxEmployees_RemoveEmployee_Click;
            mnuSetLeader.Click += CtxEmployees_SetLeader_Click;
            ctxEmployees.Items.AddRange(new ToolStripItem[] { mnuAddEmployee, mnuRemoveEmployee, new ToolStripSeparator(), mnuSetLeader });
            dgvEmployees.ContextMenuStrip = ctxEmployees;

            // Row formatting for leader highlighting
            dgvEmployees.CellFormatting -= DgvEmployees_CellFormatting;
            dgvEmployees.CellFormatting += DgvEmployees_CellFormatting;
        }

        private void DgvEmployees_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            HideEmployeeIdColumn();
            dgvEmployees.Invalidate();
        }

        private void DgvEmployees_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var row = dgvEmployees.Rows[e.RowIndex];
            var employee = row.DataBoundItem as EmployeeModel;
            if (employee != null && employee.IsLeader)
            {
                e.CellStyle.BackColor = Color.LightGoldenrodYellow;
                e.CellStyle.Font = new Font(dgvEmployees.DefaultCellStyle.Font ?? dgvEmployees.Font, FontStyle.Bold);
            }
        }

        private void CtxEmployees_AddEmployee_Click(object sender, EventArgs e)
        {
            using (var frmContacts = new FormContacts(this)) { frmContacts.ShowDialog(); }
        }

        private void CtxEmployees_RemoveEmployee_Click(object sender, EventArgs e)
        {
            btnRemoveEmployee_Click(sender, e);
        }

        private void CtxEmployees_SetLeader_Click(object sender, EventArgs e)
        {
            var row = dgvEmployees.CurrentRow;
            if (row == null) return;
            var employee = row.DataBoundItem as EmployeeModel;
            if (employee == null) return;

            // Clear existing leader
            foreach (var emp in _employeesBindingList)
                emp.IsLeader = false;

            // Set new leader
            employee.IsLeader = true;

            var fullName = ($"{employee.FirstName} {employee.LastName}").Trim();
            if (string.IsNullOrWhiteSpace(fullName)) fullName = employee.EmailAddress ?? string.Empty;
            txtProjectLeader.Text = fullName;

            // Persist
            if (!string.IsNullOrEmpty(settingsModel.XmlEmployeesFile))
                settingsModel.XmlEmployeesFile.EmployeesModelToXmlFile(_employeesBindingList.ToList());

            dgvEmployees.Invalidate();
        }

        private void HideEmployeeIdColumn()
        {
            foreach (DataGridViewColumn column in dgvEmployees.Columns)
            {
                if (string.Equals(column.Name, "Id", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(column.DataPropertyName, "Id", StringComparison.OrdinalIgnoreCase))
                {
                    column.Visible = false;
                }
            }
        }

        private void ConfigureAttachmentsDataGrid()
        {
            dgvAttachments.AutoGenerateColumns = false;
            dgvAttachments.Columns.Clear();
            dgvAttachments.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn { Name = "attachmentId", DataPropertyName = "attachmentId", Visible = false },
                new DataGridViewCheckBoxColumn { Name = "isChecked", DataPropertyName = "isChecked", HeaderText = "V", TrueValue = true, FalseValue = false, ThreeState = false },
                new DataGridViewTextBoxColumn { Name = "fileName", DataPropertyName = "fileName", HeaderText = "שם קובץ" },
                new DataGridViewTextBoxColumn { Name = "fileSize", DataPropertyName = "fileSize", HeaderText = "גודל" }
            });
            dgvAttachments.ReadOnly = true;
            dgvAttachments.BackgroundColor = SystemColors.Window;
            dgvAttachments.ForeColor = SystemColors.WindowText;
            dgvAttachments.DefaultCellStyle.BackColor = SystemColors.Window;
            dgvAttachments.DefaultCellStyle.ForeColor = SystemColors.WindowText;
            dgvAttachments.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
            dgvAttachments.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;
            dgvAttachments.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.Control;
            dgvAttachments.AlternatingRowsDefaultCellStyle.ForeColor = SystemColors.ControlText;
            dgvAttachments.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.Control;
            dgvAttachments.ColumnHeadersDefaultCellStyle.ForeColor = SystemColors.ControlText;
            dgvAttachments.EnableHeadersVisualStyles = false;
            dgvAttachments.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvAttachments.AllowUserToAddRows = false;
            dgvAttachments.RowHeadersVisible = false;
            dgvAttachments.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;

            if (dgvAttachments.Columns.Count > 1)
            {
                dgvAttachments.Columns[1].FillWeight = 10;
            }
            if (dgvAttachments.Columns.Count > 2)
            {
                dgvAttachments.Columns[2].FillWeight = 70;
            }
            if (dgvAttachments.Columns.Count > 3)
            {
                dgvAttachments.Columns[3].FillWeight = 20;
            }
        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            using (var frm = new FormSettings(this)) { frm.ShowDialog(); }
        }

        private void btnNewProject_Click(object sender, EventArgs e)
        {
            using (var frm = new FormNewProject(this)) { frm.ShowDialog(); }
        }

        private void chbOpenPDF_CheckedChanged(object sender, EventArgs e) => settingsModel.OpenPDF = chbOpenPDF.Checked;

        private void CmbSaveLocation_PathConfirmed(object sender, PathConfirmedEventArgs e)
        {
        }

        private void MainFormTaskPaneControl_Load(object sender, EventArgs e)
        {
            SetupContextMenus();
            chkbSelectAllAttachments.Checked = true;
            chkbSelectAllAttachments.Text = "הסר הכל";
            ConfigureProjectIdAutoComplete();
            LoadSearchHistory();
            settingsModel = SettingsHelpers.LoadProjectSettings();
            chkbSendNote.Checked = settingsModel.SendNoteToLeader;
            txtSubject.Text = LoadEmailSubject();
            PopulateDriveNodes();
        }
    }
}
