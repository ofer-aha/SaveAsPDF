# Changelog

All notable changes to SaveAsPDF will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2024-01-15

### Added

#### Core Features
- **Email to PDF Conversion**: Full support for converting Outlook email messages to PDF format
- **Outlook Task Pane Integration**: Modern sidebar UI with right-to-left (RTL) layout support for Hebrew
- **Project Management System**: 
  - Automatic project ID detection from email subjects
  - Project creation and configuration interface
  - XML-based project persistence
  - Default save location management
- **Attachment Management**:
  - Selective attachment extraction from emails
  - Checkbox-based attachment selection in DataGridView
  - Support for multiple file formats
- **Employee Directory**:
  - Team member management with role assignment
  - Project leader designation
  - Contact information storage (name, email)
  - Employee list persistence via XML
- **Notes Management**:
  - Dual note system: mail-specific and project-specific notes
  - Sub-tabbed interface for notes organization
  - Rich text formatting with font customization
  - Note-to-note copying functionality
- **Folder Browser**:
  - Hierarchical folder tree view
  - Drive enumeration
  - Lazy-loaded folder expansion
  - Right-click context menu support
  - Folder selection for save location

#### UI/UX Enhancements
- **Dark Mode Support**: Automatic Windows system theme detection
  - Uses `SystemColors.Window` and `SystemColors.Control` for theme adaptation
  - Full dark mode compatibility on Windows 10+ and Windows 11
- **Right-to-Left (RTL) Layout**: Complete RTL support for Hebrew and Arabic languages
  - All controls properly mirrored
  - Tab alignment set to right
  - Flow direction adjusted for RTL languages
- **Tabbed Interface Redesign**:
  - New main tab structure: Attachments, Employees, Folders, Notes
  - Notes tab with sub-tabs for mail and project notes
  - `TabAlignment.Right` for vertical tab display
  - Tab order optimization for intuitive navigation
- **Color Contrast Improvements**:
  - High contrast DataGridView styling
  - Proper text-to-background color ratios in light and dark modes
  - Alternating row colors for better readability
  - Selection color highlighting
- **Status Strip Help System**: Contextual help messages on control hover
- **Ribbon Interface**: Quick access buttons in Outlook ribbon

#### Technical Implementation
- **Platform**: .NET Framework 4.7.2 with C# 7.3 features
- **Office Integration**: Microsoft Office Interop libraries (Outlook, Word)
- **Windows Forms**: Modern form-based UI design
- **Data Persistence**:
  - XML-based configuration storage
  - JSON serialization with Newtonsoft.Json
  - Auto-save functionality for project data
- **Helper Utilities**:
  - `MailLogic.cs`: Email processing and PDF conversion logic
  - `SettingsHelpers.cs`: Application settings management
  - `XmlFileHelper.cs`: XML serialization/deserialization
  - `TreeHelpers.cs`: Folder tree operations
  - `ContextMenuHelpers.cs`: Context menu setup
  - `ComboBoxExtensions.cs`: ComboBox customization
  - `TextHelpers.cs`: Text processing utilities
  - `FileFoldersHelper.cs`: File and folder operations
  - `OfficeHelpers.cs`: Office integration utilities
  - `HtmlHelper.cs`: HTML processing
  - `JsonHelper.cs`: JSON operations

#### Data Models
- **ProjectModel**: Stores project metadata (number, name, notes, save paths)
- **EmployeeModel**: Team member information with leadership flag
- **SettingsModel**: Application configuration and user preferences
- **AttachmentsModel**: Email attachment metadata
- **DocumentModel**: Converted document information

#### Configuration Management
- XML configuration files: `Project.xml`, `Employees.xml`, `Settings.xml`
- Default folder structure template: `DefaultTree.fld`
- Application settings persistence via .NET Settings framework
- Search history with auto-complete functionality

#### Dialogs and Forms
- **FormMain**: Main application form for standalone use
- **FormSettings**: Application preferences and configuration
- **FormContacts**: Employee phone directory
- **FormNewProject**: New project creation wizard
- **XMessageBox**: Custom message box for localized feedback
- **FolderPicker**: Native Windows folder selection dialog

#### Testing Infrastructure
- Unit test project: `SaveAsPDF.Tests`
- Benchmark suite: `BenchmarkSuite1`
- MSTest framework integration

#### Documentation
- Comprehensive README.md with:
  - Feature overview
  - Installation instructions
  - Quick start guide
  - Architecture documentation
  - Development guide
  - Troubleshooting section
  - Contributing guidelines

### Changed

#### Task Pane Control Improvements
- **MainFormTaskPaneControl.cs** refactored with:
  - Improved layout using nested TableLayoutPanel and FlowLayoutPanel
  - System color theme integration throughout
  - Tab order reorganization (Attachments, Employees, Folders, Notes)
  - Enhanced color contrast for DataGridViews
  - Better responsive layout management
  - Fixed empty data source issue in dgvAttachments
  - AllowUserToAddRows set to false for proper grid initialization

#### DataGridView Configuration
- **dgvAttachments**:
  - Added checkbox column for attachment selection
  - Implemented proper column configuration
  - Fixed IndexOutOfRangeException on control creation
  - Applied system color styling for light/dark modes
  - AutoSizeColumnsMode set to Fill
- **dgvEmployees**:
  - Configured columns: Id (hidden), FirstName, LastName, EmailAddress
  - Applied system color styling
  - Proper data binding with BindingList

#### Color Scheme Refactoring
- Replaced hardcoded colors with SystemColors throughout
- Implemented `ApplySystemColors()` method for base control coloring
- Updated all form controls to use theme-aware colors:
  - Labels, TextBoxes, ComboBoxes, CheckBoxes, Buttons, RichTextBoxes
  - DataGridViews with alternating row styles
  - TreeView and TabControl styling

#### Tab Structure Reorganization
- Moved Notes tab to last position (previously first)
- Adjusted tab alignment to `TabAlignment.Right` for vertical display
- Implemented sub-TabControl for Notes with two sub-tabs:
  - "הערות למייל" (Mail Notes)
  - "הערות בפרויקט" (Project Notes)
- Updated main tab order: Attachments → Employees → Folders → Notes

#### Interface Styling
- All buttons now use `UseVisualStyleBackColor = true`
- Consistent spacing and layout management
- Improved panel organization with appropriate BackColor assignments
- Better visual hierarchy through color differentiation

### Fixed

#### Critical Issues
- **DataGridView Initialization Error**: Fixed `IndexOutOfRangeException: Index -1 does not have a value`
  - Root cause: DataSource bound to empty collection during handle creation
  - Solution: Deferred DataSource assignment until data is populated
  - Affected: dgvAttachments
  
#### Color/Theme Issues
- **Missing Dark Mode Support**: Now properly adapts to Windows system theme
- **Poor Text Contrast**: Fixed by implementing proper SystemColors usage
- **Tab Visibility**: Improved tab clarity with consistent styling

#### Layout Issues
- **Notes Buttons Placement**: Moved to proper position after tab control
- **Empty Grid Error**: Prevented by deferring data binding in ConfigureAttachmentsDataGrid()
- **Control Initialization**: Proper BackColor and ForeColor assignment from start

### Improved

#### Code Quality
- Better separation of concerns in layout configuration
- Improved method organization in BuildLayout()
- More consistent naming conventions
- Enhanced error handling in data loading methods

#### User Experience
- Clearer visual hierarchy
- Better color contrast for accessibility
- More intuitive tab organization
- Responsive layout that adapts to system theme
- Improved button labeling and positioning

#### Documentation
- Added comprehensive inline comments for complex logic
- Detailed docstrings for public methods
- README.md with full feature documentation
- Code style guidelines for contributors

### Known Limitations

- Email body formatting may vary in PDF output depending on HTML complexity
- Very large attachments may impact performance
- Batch email processing not yet implemented
- Cloud storage integration pending (OneDrive, SharePoint)

## [0.9.0] - 2023-12-01

### Initial Development
- Project scaffolding and setup
- Basic Outlook add-in integration
- Form layouts and controls initialization
- Core helper classes implementation

## Migration Guide

### From Pre-1.0 to 1.0.0

If you're upgrading from earlier versions, note the following changes:

1. **Tab Order**: The Notes tab has moved to the last position
   - Update any documentation or user training materials
   - Search functionality still works the same way

2. **UI Theme Support**: 
   - Dark mode is now fully supported
   - No action needed - it works automatically
   - Custom colors will be overridden by system theme

3. **DataGridView Access**:
   - Ensure empty collections are handled gracefully
   - Check any custom code that directly accesses DataGridView rows

4. **Configuration Files**:
   - Existing XML files remain compatible
   - New settings will use SystemColors instead of hardcoded values

## Future Roadmap

### Planned for v1.1.0
- [ ] Batch email conversion
- [ ] Email template customization
- [ ] Enhanced project statistics and reporting
- [ ] Improved attachment preview

### Planned for v1.2.0
- [ ] Cloud storage integration (OneDrive, SharePoint)
- [ ] Advanced search and filtering
- [ ] Email archiving and tagging
- [ ] Multi-language localization (beyond Hebrew/Arabic)

### Planned for v2.0.0
- [ ] Modern .NET 6+ migration
- [ ] WPF UI redesign
- [ ] Enhanced performance optimization
- [ ] Mobile companion app
- [ ] REST API for external integrations

## Security

No security issues discovered in v1.0.0. All user data is stored locally in XML files.

## Dependencies

### NuGet Packages
- **Microsoft Office Interop Outlook**: 15.0.0.0 - Office automation
- **Newtonsoft.Json**: 13.0.3 - JSON serialization
- **Ookii.Dialogs.WinForms**: 4.0.0 - Native Windows dialogs
- **Microsoft.Extensions.* (various)**: 3.1.0 - Dependency injection and configuration

### Development Dependencies
- **MSTest.TestFramework**: 2.2.10 - Unit testing
- **MSTest.TestAdapter**: 2.2.10 - Test execution

## Contributors

- Initial development team
- Community testers and feedback providers
- Microsoft Office documentation contributors

## Support

For bugs, feature requests, or questions:
- **GitHub Issues**: [Create an issue](https://github.com/ofer-aha/SaveAsPDF/issues)
- **Discussions**: [Start a discussion](https://github.com/ofer-aha/SaveAsPDF/discussions)

---

**Version History**:
- v1.0.0 - 2024-01-15 (Current)
- v0.9.0 - 2023-12-01 (Initial Development)

For more information, visit the [SaveAsPDF GitHub Repository](https://github.com/ofer-aha/SaveAsPDF).
