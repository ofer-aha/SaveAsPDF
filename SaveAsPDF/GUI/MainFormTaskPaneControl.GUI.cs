using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;

namespace SaveAsPDF
{
    /// <summary>
    /// GUI partial class for MainFormTaskPaneControl.
    /// Contains all UI layout construction and visual element initialization.
    /// This file is separated from business logic to maintain clean separation of concerns.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("7F5A3B2C-9E1D-4D6F-B8C4-1A2E3D4F5678")]
    [ProgId("SaveAsPDF.MainFormTaskPaneControl")]
    public partial class MainFormTaskPaneControl
    {
        /// <summary>
        /// Initializes the task pane control and constructs the complete UI layout.
        /// Sets up the control hierarchy, event handlers, and applies system theming.
        /// </summary>
        public MainFormTaskPaneControl()
        {
            // Configure base UserControl properties
            Dock = DockStyle.Fill;              // Fill entire parent container
            AutoScroll = true;                  // Enable scrolling if content exceeds viewport
            RightToLeft = RightToLeft.Yes;      // Set RTL layout for Hebrew text

            // Apply theme-aware system colors
            ApplySystemColors();

            // Build the complete control hierarchy
            BuildLayout();

            // Configure data grids with columns and styling
            ConfigureEmployeeDataGrid();
            ConfigureAttachmentsDataGrid();

            // Wire up event handlers
            KeyDown += MainFormTaskPaneControl_KeyDown;
            btnProjectLeader.Click += btnProjectLeader_Click;
            
            // Disable validation on settings to allow immediate action
            btnSettings.CausesValidation = false;
            
            // Wire up the Load event for initialization tasks
            Load += MainFormTaskPaneControl_Load;
        }

        /// <summary>
        /// Constructs the complete UI layout hierarchy using nested TableLayoutPanels.
        /// Layout structure:
        ///   Main (3 rows)
        ///     ├─ Top: GroupBox with project/email metadata fields
        ///     ├─ Middle: TabControl with attachments, employees, folders, notes
        ///     └─ Bottom: Action buttons and status strip
        /// </summary>
        private void BuildLayout()
        {
            // === ROOT LAYOUT PANEL ===
            // Main container with 3 rows: top metadata, middle tabs, bottom actions
            var main = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                RightToLeft = RightToLeft.Yes
            };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));    // Top: auto-size for metadata
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Middle: fill remaining space
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));    // Bottom: auto-size for buttons

            // === TOP SECTION: PROJECT AND EMAIL METADATA ===
            var topGroup = new GroupBox
            {
                Text = "נתוני פרויקט והודעה",  // "Project and Message Data"
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };

            // Table for label-field pairs (2 columns: labels 30%, fields 70%)
            var topTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control
            };
            topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            topTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));

            // ROW 0: Project ID field with validation
            var lblProjectID = new Label
            {
                Text = "מספר פרויקט",  // "Project Number"
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            txtProjectID.Dock = DockStyle.Fill;
            txtProjectID.BackColor = SystemColors.Window;
            txtProjectID.ForeColor = SystemColors.WindowText;
            txtProjectID.RightToLeft = RightToLeft.No;
            txtProjectID.TextAlign = HorizontalAlignment.Left;
            txtProjectID.KeyDown += txtProjectID_KeyDown;           // Enter key handler
            txtProjectID.Validating += txtProjectID_Validating;     // Format validation
            txtProjectID.Validated += txtProjectID_Validated;       // Load project after validation
            topTable.Controls.Add(lblProjectID, 0, 0);
            topTable.Controls.Add(txtProjectID, 1, 0);

            // ROW 1: Project Name (read-only, populated from project data)
            var lblProjectName = new Label
            {
                Text = "שם הפרויקט",  // "Project Name"
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            txtProjectName.Dock = DockStyle.Fill;
            txtProjectName.BackColor = SystemColors.Window;
            txtProjectName.ForeColor = SystemColors.WindowText;
            topTable.Controls.Add(lblProjectName, 0, 1);
            topTable.Controls.Add(txtProjectName, 1, 1);

            // ROW 2: Project Leader (textbox + button to select from contacts)
            var lblLeader = new Label
            {
                Text = "מתכנן מוביל",  // "Lead Planner"
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            var pnlLeader = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Fill,
                AutoSize = true,
                BackColor = SystemColors.Control
            };
            txtProjectLeader.Width = 160;
            txtProjectLeader.BackColor = SystemColors.Window;
            txtProjectLeader.ForeColor = SystemColors.WindowText;
            btnProjectLeader.Text = "בחר מתכנן מוביל";
            btnProjectLeader.Width = 100;
            btnProjectLeader.BackColor = SystemColors.Control;
            btnProjectLeader.ForeColor = SystemColors.ControlText;
            btnProjectLeader.UseVisualStyleBackColor = true;
            pnlLeader.Controls.Add(txtProjectLeader);
            pnlLeader.Controls.Add(btnProjectLeader);
            topTable.Controls.Add(lblLeader, 0, 2);
            topTable.Controls.Add(pnlLeader, 1, 2);

            // ROW 3: Email Subject (auto-populated from selected mail item)
            var lblSubject = new Label
            {
                Text = "נושא ההודעה",  // "Message Subject"
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            txtSubject.Dock = DockStyle.Fill;
            txtSubject.BackColor = SystemColors.Window;
            txtSubject.ForeColor = SystemColors.WindowText;
            topTable.Controls.Add(lblSubject, 0, 3);
            topTable.Controls.Add(txtSubject, 1, 3);

            // ROW 4: Save Location (Explorer-style address bar + folder picker button)
            var lblSaveLocation = new Label
            {
                Text = "מיקום שמירה",  // "Save Location"
                AutoSize = true,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText
            };
            var pnlSaveLocation = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Fill,
                AutoSize = true,
                BackColor = SystemColors.Control,
                RightToLeft = RightToLeft.No                        // LTR for address bar
            };
            cmbSaveLocation.Width = 280;
            cmbSaveLocation.BackColor = SystemColors.Window;
            cmbSaveLocation.ForeColor = SystemColors.WindowText;
            cmbSaveLocation.RightToLeft = RightToLeft.No;
            cmbSaveLocation.PathConfirmed += CmbSaveLocation_PathConfirmed;  // Update breadcrumb on path change
            btnFolders.Text = "בחר תיקייה";  // "Select Folder"
            btnFolders.BackColor = SystemColors.Control;
            btnFolders.ForeColor = SystemColors.ControlText;
            btnFolders.UseVisualStyleBackColor = true;
            btnFolders.Click += btnFolders_Click;
            pnlSaveLocation.Controls.Add(cmbSaveLocation);
            pnlSaveLocation.Controls.Add(btnFolders);
            topTable.Controls.Add(lblSaveLocation, 0, 4);
            topTable.Controls.Add(pnlSaveLocation, 1, 4);

            topGroup.Controls.Add(topTable);

            // === MIDDLE SECTION: TABBED INTERFACE ===
            // Main tab control with 4 tabs: Attachments, Employees, Folders, Notes
            var mainTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes,
                Alignment = TabAlignment.Right                      // Tabs on right for RTL
            };

            // --- TAB: NOTES (contains nested sub-tabs for mail notes and project notes) ---
            var tabNotes = new TabPage("הערות")  // "Notes"
            {
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                UseVisualStyleBackColor = true
            };
            var subTabNotes = new TabControl
            {
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes,
                Alignment = TabAlignment.Right
            };
            
            // Sub-tab: Mail Notes (now includes buttons for copying to project and font selection)
            var tabMailNotes = new TabPage("הערות למייל")  // "Mail Notes"
            {
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                UseVisualStyleBackColor = true
            };
            var mailNotesLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            mailNotesLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mailNotesLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var pnlMailNoteButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                BackColor = SystemColors.Control
            };
            btnCopyNotesToProject.Text = "העתק ממייל → פרויקט";  // "Copy from Mail → Project"
            btnCopyNotesToProject.BackColor = SystemColors.Control;
            btnCopyNotesToProject.ForeColor = SystemColors.ControlText;
            btnCopyNotesToProject.UseVisualStyleBackColor = true;
            btnCopyNotesToProject.Click += btnCopyNotesToProject_Click;
            btnStyle.Text = "גופן";  // "Font"
            btnStyle.BackColor = SystemColors.Control;
            btnStyle.ForeColor = SystemColors.ControlText;
            btnStyle.UseVisualStyleBackColor = true;
            btnStyle.Click += btnStyle_Click;
            pnlMailNoteButtons.Controls.Add(btnCopyNotesToProject);
            pnlMailNoteButtons.Controls.Add(btnStyle);
            rtxtNotes.Dock = DockStyle.Fill;
            rtxtNotes.BackColor = SystemColors.Window;
            rtxtNotes.ForeColor = SystemColors.WindowText;
            mailNotesLayout.Controls.Add(pnlMailNoteButtons, 0, 0);
            mailNotesLayout.Controls.Add(rtxtNotes, 0, 1);
            tabMailNotes.Controls.Add(mailNotesLayout);

            // --- TAB: PROJECT NOTES (with copy-to-mail button and editor) ---
            var tabProjectNotes = new TabPage("הערות בפרויקט")  // "Project Notes"
            {
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                UseVisualStyleBackColor = true
            };
            var projectNotesLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            projectNotesLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            projectNotesLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var pnlProjectNoteButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                BackColor = SystemColors.Control
            };
            btnCopyNotesToMail.Text = "העתק לפרויקט ← מייל";  // "Copy to Project ← Mail"
            btnCopyNotesToMail.BackColor = SystemColors.Control;
            btnCopyNotesToMail.ForeColor = SystemColors.ControlText;
            btnCopyNotesToMail.UseVisualStyleBackColor = true;
            btnCopyNotesToMail.Click += btnCopyNotesToMail_Click;
            pnlProjectNoteButtons.Controls.Add(btnCopyNotesToMail);
            rtxtProjectNotes.Dock = DockStyle.Fill;
            rtxtProjectNotes.BackColor = SystemColors.Window;
            rtxtProjectNotes.ForeColor = SystemColors.WindowText;
            projectNotesLayout.Controls.Add(pnlProjectNoteButtons, 0, 0);
            projectNotesLayout.Controls.Add(rtxtProjectNotes, 0, 1);
            tabProjectNotes.Controls.Add(projectNotesLayout);

            subTabNotes.TabPages.Add(tabProjectNotes);
            subTabNotes.TabPages.Add(tabMailNotes);
            tabNotes.Controls.Add(subTabNotes);

            // --- TAB: ATTACHMENTS (with select-all checkbox and data grid) ---
            var tabAttachments = new TabPage("קבצים מצורפים")  // "Attachments"
            {
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                UseVisualStyleBackColor = true
            };
            var attachmentsTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                BackColor = SystemColors.Control
            };

            attachmentsTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            attachmentsTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            
            // Checkbox to toggle all attachment selections
            chkbSelectAllAttachments.Text = "הסר הכל";  // "Remove All"
            chkbSelectAllAttachments.Checked = true;
            chkbSelectAllAttachments.BackColor = SystemColors.Control;
            chkbSelectAllAttachments.ForeColor = SystemColors.ControlText;
            chkbSelectAllAttachments.CheckedChanged += chkbSelectAllAttachments_CheckedChanged;
            attachmentsTable.Controls.Add(chkbSelectAllAttachments, 0, 0);
            
            // Data grid showing attachment list
            dgvAttachments.Dock = DockStyle.Fill;
            dgvAttachments.CellDoubleClick += dgvAttachments_CellDoubleClick;
            attachmentsTable.Controls.Add(dgvAttachments, 0, 1);
            tabAttachments.Controls.Add(attachmentsTable);

            // --- TAB: EMPLOYEES (with phonebook and remove buttons) ---
            var tabEmployees = new TabPage("עובדים בפרויקט")  // "Project Employees"
            {
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                UseVisualStyleBackColor = true
            };
            var employeesTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                BackColor = SystemColors.Control
            };
            employeesTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            employeesTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            
            // Button panel for employee management
            var pnlEmpButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                BackColor = SystemColors.Control
            };
            btnPhoneBook.Text = "ספר טלפונים";  // "Phone Book"
            btnPhoneBook.BackColor = SystemColors.Control;
            btnPhoneBook.ForeColor = SystemColors.ControlText;
            btnPhoneBook.UseVisualStyleBackColor = true;
            btnPhoneBook.Click += btnPhoneBook_Click;
            btnRemoveEmployee.Text = "הסר עובד";  // "Remove Employee"
            btnRemoveEmployee.BackColor = SystemColors.Control;
            btnRemoveEmployee.ForeColor = SystemColors.ControlText;
            btnRemoveEmployee.UseVisualStyleBackColor = true;
            btnRemoveEmployee.Click += btnRemoveEmployee_Click;
            pnlEmpButtons.Controls.Add(btnPhoneBook);
            pnlEmpButtons.Controls.Add(btnRemoveEmployee);
            employeesTable.Controls.Add(pnlEmpButtons, 0, 0);
            
            // Data grid showing employee list
            dgvEmployees.Dock = DockStyle.Fill;
            dgvEmployees.CurrentCellDirtyStateChanged += dgvEmployees_CurrentCellDirtyStateChanged;
            dgvEmployees.CellValueChanged += dgvEmployees_CellValueChanged;
            employeesTable.Controls.Add(dgvEmployees, 0, 1);
            tabEmployees.Controls.Add(employeesTable);

            // --- TAB: FOLDERS (tree view of project folder structure) ---
            var tabFolders = new TabPage("עץ תיקיות פרויקט")  // "Project Folder Tree"
            {
                RightToLeft = RightToLeft.Yes,
                BackColor = SystemColors.Control,
                UseVisualStyleBackColor = true
            };
            tvFolders.Dock = DockStyle.Fill;
            tvFolders.BackColor = SystemColors.Window;
            tvFolders.ForeColor = SystemColors.WindowText;
            tvFolders.PathSeparator = "\\";
            tvFolders.BeforeExpand += tvFolders_BeforeExpand;               // Lazy-load child nodes
            tvFolders.BeforeCollapse += tvFolders_BeforeCollapse;
            tvFolders.MouseDown += tvFolders_MouseDown;
            tvFolders.NodeMouseClick += tvFolders_NodeMouseClick;           // Select path on click
            tvFolders.NodeMouseDoubleClick += tvFolders_NodeMouseDoubleClick;  // Open in Explorer
            tabFolders.Controls.Add(tvFolders);

            // Add all tabs to main tab control
            mainTabControl.TabPages.Add(tabAttachments);
            mainTabControl.TabPages.Add(tabEmployees);
            mainTabControl.TabPages.Add(tabFolders);
            mainTabControl.TabPages.Add(tabNotes);

            // Wrapper panel for tabs only (buttons moved into Project Notes tab)
            var middlePanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 1,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = SystemColors.Control
            };
            middlePanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            middlePanel.Controls.Add(mainTabControl, 0, 0);

            // Store reference to main tab control for later access
            this.tabNotes = mainTabControl;

            // === BOTTOM SECTION: ACTION BUTTONS AND STATUS BAR ===
            var bottomPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = SystemColors.Control
            };
            bottomPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // Row 0: checkboxes
            bottomPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // Row 1: buttons
            bottomPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // Row 2: status strip

            // --- Row 0: Checkboxes, each on its own line ---
            var pnlCheckboxes = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                BackColor = SystemColors.Control
            };

            // Option to open PDF after saving
            chbOpenPDF.Text = "פתח PDF לאחר שמירה";  // "Open PDF after saving"
            chbOpenPDF.AutoSize = true;
            chbOpenPDF.BackColor = SystemColors.Control;
            chbOpenPDF.ForeColor = SystemColors.ControlText;
            chbOpenPDF.Margin = new Padding(3, 2, 3, 0);
            chbOpenPDF.CheckedChanged += chbOpenPDF_CheckedChanged;

            // Option to send note to project leader after saving
            chkbSendNote.Text = "שלח לראש הפרויקט";  // "Send to project leader"
            chkbSendNote.AutoSize = true;
            chkbSendNote.BackColor = SystemColors.Control;
            chkbSendNote.ForeColor = SystemColors.ControlText;
            chkbSendNote.Margin = new Padding(3, 2, 3, 0);
            chkbSendNote.Checked = false;  // Default off; overridden from settings in Load

            pnlCheckboxes.Controls.Add(chbOpenPDF);
            pnlCheckboxes.Controls.Add(chkbSendNote);

            // --- Row 1: Action buttons ---
            var pnlButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true,
                BackColor = SystemColors.Control
            };

            // Primary action: Save to PDF
            btnOK.Text = "שמור ל-PDF";  // "Save to PDF"
            btnOK.AutoSize = true;
            btnOK.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            btnOK.Padding = new Padding(4, 0, 4, 0);
            btnOK.BackColor = SystemColors.Control;
            btnOK.ForeColor = SystemColors.ControlText;
            btnOK.UseVisualStyleBackColor = true;
            btnOK.Click += btnOK_Click;

            // Settings dialog
            btnSettings.Text = "הגדרות";  // "Settings"
            btnSettings.AutoSize = true;
            btnSettings.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            btnSettings.Padding = new Padding(4, 0, 4, 0);
            btnSettings.BackColor = SystemColors.Control;
            btnSettings.ForeColor = SystemColors.ControlText;
            btnSettings.UseVisualStyleBackColor = true;
            btnSettings.Click += BtnSettings_Click;

            // New project creation
            btnNewProject.Text = "פרויקט חדש";  // "New Project"
            btnNewProject.AutoSize = true;
            btnNewProject.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            btnNewProject.Padding = new Padding(4, 0, 4, 0);
            btnNewProject.BackColor = SystemColors.Control;
            btnNewProject.ForeColor = SystemColors.ControlText;
            btnNewProject.UseVisualStyleBackColor = true;
            btnNewProject.Click += btnNewProject_Click;

            pnlButtons.Controls.Add(btnOK);
            pnlButtons.Controls.Add(btnSettings);
            pnlButtons.Controls.Add(btnNewProject);

            // Status strip for displaying contextual help and validation messages
            statusStrip.Items.Add(tsslStatus);
            statusStrip.RightToLeft = RightToLeft.Yes;
            statusStrip.BackColor = SystemColors.Control;
            statusStrip.ForeColor = SystemColors.ControlText;

            bottomPanel.Controls.Add(pnlCheckboxes, 0, 0);
            bottomPanel.Controls.Add(pnlButtons, 0, 1);
            bottomPanel.Controls.Add(statusStrip, 0, 2);

            // === ASSEMBLE FINAL LAYOUT ===
            main.Controls.Add(topGroup, 0, 0);      // Top: metadata fields
            main.Controls.Add(middlePanel, 0, 1);   // Middle: tabs
            main.Controls.Add(bottomPanel, 0, 2);   // Bottom: buttons + status
            Controls.Add(main);
            
            // Configure error provider for validation feedback
            errorProviderMain.ContainerControl = this;
            errorProviderMain.RightToLeft = true;
            
            // Wire up mouse hover status help for all controls
            WireStatusHelp();
        }
    }
}
