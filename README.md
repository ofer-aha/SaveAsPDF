# SaveAsPDF - Outlook Add-In

> ⚠️ **WORK IN PROGRESS** — This project is under active development. There is currently **no installation package or release available**. The only way to run the add-in is by building from source (see [Building from Source](#building-from-source)).

Convert Outlook mail messages to PDF files with ease. SaveAsPDF is a Microsoft Outlook COM add-in that enables users to save emails and their attachments as PDF documents, with full support for project management, employee tracking, and advanced customization options.

## Features

- **📧 Email to PDF Conversion**: Convert any Outlook email message to a professional PDF file
- **📎 Attachment Management**: Select and save email attachments alongside the PDF
- **🏢 Project Management**: Organize emails by project with automatic project ID detection from email subjects
- **👥 Employee Directory**: Manage project team members with contact information and role assignment
- **📝 Rich Notes**: Add and maintain both mail-specific and project-specific notes
- **🌳 Folder Structure**: Browse and manage save locations with a Windows Explorer-style breadcrumb address bar and folder tree view
- **🎨 Customization**: Configure fonts, colors, and application settings to your preference
- **🌍 Right-to-Left Support**: Full support for Hebrew and other RTL languages
- **🌙 Dark Mode Support**: Automatic theme adaptation to Windows system colors (light/dark modes)
- **⌨️ Quick Access**: Ribbon interface for quick conversion of selected messages
- **🔒 Built-in Security**: Input validation, path sanitization, HTML encoding, and safe file naming throughout

## Security Features

Security is built into the core of SaveAsPDF rather than added as an afterthought. The following protections are implemented and active:

### Input Validation
- **Project ID validation** (`SafeProjectID`): All project IDs are validated against a strict regex pattern before being used in any file system operation. Invalid IDs are rejected at the UI and helper layers.
- **Folder name sanitization** (`SafeFolderName`): Every folder and file name derived from user input or email data is scrubbed of all characters that are invalid on Windows file systems. Windows reserved device names (`CON`, `PRN`, `AUX`, `NUL`, `COM1–COM3`, `LPT1–LPT3`) are automatically renamed to avoid collisions.
- **Path normalization**: All paths are resolved to their canonical form using `Path.GetFullPath` before use, preventing path traversal attacks via relative path components (e.g. `..`).

### HTML Output Security
- **HTML encoding**: Every piece of user-supplied or email-derived data written into generated HTML (project name, project ID, notes, user name, employee names, email addresses, attachment names) is HTML-encoded with `WebUtility.HtmlEncode` before insertion, preventing cross-site scripting (XSS) in Outlook's HTML renderer.
- **URI encoding**: File paths embedded as `file:///` links in generated HTML have special URI characters (`space`, `#`, `?`, `%`) percent-encoded, preventing link injection while correctly preserving non-ASCII (Hebrew/Arabic) characters.

### File System Safety
- **Unique file name generation**: If a target file or folder already exists, a numbered suffix is appended automatically, preventing silent overwrites.
- **Safe attachment saving**: Attachment file names are sanitized by stripping all invalid file-system characters before being written to disk.
- **Password-protected document detection**: Word documents that are password-protected are detected before conversion is attempted, avoiding hangs or errors.
- **Hidden configuration directories**: The `.SaveAsPDF` metadata folder is created as a hidden directory, reducing accidental user modification or deletion of project data.
- **File attribute handling**: Read-only or hidden attributes on existing XML configuration files are cleared before overwriting, preventing silent write failures.

### Error Handling and Access Control
- **`UnauthorizedAccessException` handling**: File and directory operations explicitly catch access-denied errors and surface them to the user without crashing the add-in.
- **COM object lifecycle management**: All Outlook COM objects (folders, items, stores) are released in `finally` blocks using `Marshal.ReleaseComObject`, preventing use-after-free and memory leaks.
- **Diagnostic logging**: Startup and selection-change exceptions are logged to `C:\Temp\SaveAsPDF_error.txt` for post-mortem debugging without exposing sensitive data in the UI.
- **MAPI filter injection prevention**: Contact search queries escape single-quotes before being embedded in MAPI restriction filter strings, preventing filter injection.

### Data Storage
- All user data (project metadata, employee lists, settings) is stored **locally** on the user's machine in XML files. No data is transmitted to remote servers.
- XML serialization uses `XmlReader`/`XmlWriter` with explicit `UTF-8` encoding, preventing encoding-related deserialization issues.

## System Requirements

- **Microsoft Outlook**: 2016 or later (Office 365 / Microsoft 365 compatible)
- **.NET Framework**: 4.7.2 or higher
- **Windows**: Windows 7 or later (32-bit or 64-bit)
- **RAM**: Minimum 512 MB (1 GB or more recommended)

## Installation

> ⚠️ **No installer is available yet.** This project is a work in progress. You must build the add-in from source and register it manually (see below).

### Building from Source

1. **Clone the repository**:
   ```bash
   git clone https://github.com/ofer-aha/SaveAsPDF.git
   cd SaveAsPDF
   ```

2. **Open `SaveAsPDF.sln`** in Visual Studio 2019 or later.

3. **Restore NuGet packages**:
   ```bash
   nuget restore SaveAsPDF.sln
   ```

4. **Build the solution** (Release configuration recommended):
   ```bash
   msbuild SaveAsPDF.sln /p:Configuration=Release
   ```

5. **Register the COM add-in** (run as Administrator):
   ```bash
   regasm /codebase SaveAsPDF\bin\Release\SaveAsPDF.dll
   ```
   This registers the add-in under `HKCU\Software\Microsoft\Office\Outlook\Addins\SaveAsPDF.Connect` with `LoadBehavior=3` (load at Outlook startup).

6. **Restart Microsoft Outlook** — the SaveAsPDF ribbon button will appear.

### Run Tests

```bash
dotnet test SaveAsPDF.Tests\SaveAsPDF.Tests.csproj
```

## Quick Start

### Converting an Email to PDF

1. **Select an Email**: Click on the email you want to convert in Outlook.
2. **Click "Save as PDF"**: Use the SaveAsPDF button in the Outlook ribbon to open the task pane.
3. **Configure Options**:
   - Enter or select a project number (auto-detected from the email subject when possible)
   - Choose attachments to include
   - Add notes or project-specific information
   - Select the save location using the breadcrumb address bar or folder tree
4. **Save**: Click the "Save as PDF" button to generate the file.

### Managing Projects

1. **Create a New Project**: Click "New Project" in the task pane.
   - Set project number, name, and default save location.
   - Add team member information.

2. **Add Employees**: Access the Employees tab to manage team member details and assign the project leader role.

3. **View Project Notes**: Switch between mail notes and project notes using the sub-tabs.

### Project ID Auto-Detection

Email subjects with project IDs are automatically recognised. For example:
- Subject: `Project 1234 - Meeting Notes` → detected project ID: `1234`
- Supported formats: 3–4 digit IDs with optional letter suffix and up to two hyphenated sub-segments (e.g. `1234`, `1234A`, `1234-1`, `1234-1-2`)

## Architecture

### Project Structure

```
SaveAsPDF/
├── SaveAsPDF/                      # Main add-in project (COM add-in, no VSTO runtime)
│   ├── GUI_Logic/                  # Task pane business logic
│   │   └── MainFormTaskPaneControl.cs
│   ├── GUI/                        # Task pane layout (designer-generated)
│   │   └── MainFormTaskPaneControl.GUI.cs
│   ├── Helpers/                    # Helper classes
│   │   ├── FileFoldersHelper.cs    # Path sanitization, safe folder/file names
│   │   ├── HtmlHelper.cs           # HTML generation with XSS protection
│   │   ├── MailLogic.cs            # Email sending utilities
│   │   ├── OfficeHelpers.cs        # Word/Outlook COM interop, PDF conversion
│   │   ├── OutlookProcessor.cs     # Outlook contact and folder enumeration
│   │   ├── PathBreadcrumbHelper.cs # Breadcrumb ↔ path conversion
│   │   ├── SettingsHelpers.cs      # Application settings management
│   │   ├── TextHelpers.cs          # Text processing utilities
│   │   ├── TreeHelpers.cs          # Folder tree operations
│   │   ├── XmlFileHelper.cs        # XML serialization/deserialization
│   │   └── DiagnosticsAndBehavior.cs # Logging and behavior helpers
│   ├── Models/                     # Data models
│   │   ├── ProjectModel.cs
│   │   ├── EmployeeModel.cs
│   │   ├── SettingsModel.cs
│   │   ├── AttachmentsModel.cs
│   │   ├── DocumentModel.cs
│   │   └── ContactFolderInfo.cs
│   ├── Connect.cs                  # COM add-in entry point (IDTExtensibility2)
│   ├── AddinGlobals.cs             # Global add-in state
│   ├── FormSettings.cs             # Settings dialog
│   ├── FormContacts.cs             # Employee management dialog
│   ├── FormNewProject.cs           # New project creation dialog
│   ├── SettingsTaskPaneControl.cs  # Settings task pane
│   └── WordConvert.cs              # Word-to-PDF conversion
├── SaveAsPDF.Tests/                # Unit tests (MSTest)
├── BenchmarkSuite1/                # Performance benchmarks
└── README.md                       # This file
```

### Technology Stack

| Component | Technology |
|-----------|-----------|
| Language | C# 7.3 |
| Runtime | .NET Framework 4.7.2 |
| Add-in type | COM Add-in (IDTExtensibility2) — no VSTO runtime required |
| Office integration | Microsoft Office Interop (Outlook, Word) |
| UI | Windows Forms |
| Configuration | XML (project data, employee lists, settings) |
| JSON | Newtonsoft.Json 13.0.3 |
| Dialogs | Ookii.Dialogs.WinForms 4.0.0 |
| Testing | MSTest 2.2.10 |

### Key Components

- **`Connect`**: COM add-in entry point implementing `IDTExtensibility2`, `IRibbonExtensibility`, and `ICustomTaskPaneConsumer`. Handles Outlook lifecycle, task pane creation, and ribbon callbacks.
- **`MainFormTaskPaneControl`**: Outlook sidebar task pane with RTL layout, dark mode support, tabbed interface (Attachments, Employees, Folders, Notes), and all core user interactions.
- **`FileFoldersHelper`**: `SafeProjectID` (regex-validated project IDs), `SafeFolderName` (file-system character scrubbing + reserved-name protection), path normalization, and directory management.
- **`HtmlHelper`**: Generates HTML summary files with full `WebUtility.HtmlEncode` protection on all user-supplied fields and safe `file:///` URI construction.
- **`OfficeHelpers`**: Drives Word-to-PDF conversion via Office Interop with safe attachment file name handling and full COM object lifecycle management.
- **`XmlFileHelper`**: XML serialization/deserialization for `ProjectModel`, `EmployeeModel`, and `SettingsModel`, with `UnauthorizedAccessException` handling and file attribute management.
- **`OutlookProcessor`**: Outlook contact and folder enumeration with MAPI filter quote-escaping and full COM object release in `finally` blocks.

## Configuration Files

SaveAsPDF stores project metadata in hidden files alongside the project folder:

| File | Contents |
|------|----------|
| `.SaveAsPDF_Project.xml` | Project information (name, notes, metadata) |
| `.SaveAsPDF_Emploeeys.xml` | Team member list and roles |
| App settings (`Settings.Default`) | User preferences and configurations |
| `DefaultTree.fld` | Default folder structure template |

The `.SaveAsPDF` metadata folder is created as a **hidden directory** so it does not appear during normal project file browsing.

## Interface Features

### Task Pane Layout

The Outlook sidebar provides:
- **Top Section**: Project ID and name, email subject, breadcrumb save-location selector
- **Middle Section**: Tabbed content — Attachments, Employees, Folders, Notes (with mail/project sub-tabs)
- **Bottom Section**: Action buttons (Save to PDF, Cancel, Settings, New Project), status messages, open-PDF checkbox

### Breadcrumb Address Bar

Save locations are displayed in Windows Explorer-style breadcrumb format (`C › Projects › 1000`) for readability. The control automatically converts between breadcrumb display and editable full paths on focus.

### Theme Support

- **Light Mode**: Uses `SystemColors.Window` and `SystemColors.Control`
- **Dark Mode**: Automatically adapts to the Windows system dark theme
- **RTL Layout**: Full right-to-left support for Hebrew and Arabic

## Development

### Prerequisites

- Visual Studio 2019 or later
- .NET Framework 4.7.2 Developer Pack
- Microsoft Office development tools (Outlook and Word interop assemblies)
- NuGet package manager

### Code Style

- PascalCase for class and method names; camelCase for local variables; `_` prefix for private fields
- XML documentation comments on all public members

## Troubleshooting

### Add-in Not Appearing in Outlook

1. Verify the DLL is registered: check `HKCU\Software\Microsoft\Office\Outlook\Addins\SaveAsPDF.Connect`
2. Check for disabled add-ins: File → Options → Add-ins → Manage COM Add-ins
3. Review `C:\Temp\SaveAsPDF_error.txt` for startup errors
4. Restart Outlook after registration

### PDF Conversion Issues

- **Large attachments**: Some emails may take longer to process
- **Encoding issues**: Ensure the email is in standard format (not RTF only)
- **Path issues**: Verify write permissions to the selected save location

### File Location Problems

- **UNC paths**: Network paths may require authentication
- **Special characters**: `SafeFolderName` automatically replaces invalid characters, but very long paths may still hit the Windows 260-character limit (enable long paths in Windows 10+)

## Contributing

Contributions are welcome! Please follow these guidelines:

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature/YourFeature`
3. **Make your changes**: Follow the existing code style
4. **Write tests**: Add unit tests for new functionality
5. **Submit a Pull Request**: Describe your changes and reference any issues

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

## Support

- **GitHub Issues**: [Report a bug](https://github.com/ofer-aha/SaveAsPDF/issues)
- **Discussions**: [Start a discussion](https://github.com/ofer-aha/SaveAsPDF/discussions)

## Changelog

See [SaveAsPDF/CHANGELOG.md](SaveAsPDF/CHANGELOG.md) for version history and release notes.

## Disclaimer

SaveAsPDF is provided as-is. The developers are not responsible for data loss or corruption, PDF conversion quality or accuracy, compatibility issues with specific Outlook versions, or third-party software conflicts.

## Roadmap

- [ ] Installation package / ClickOnce installer
- [ ] Cloud storage integration (OneDrive, SharePoint)
- [ ] Batch PDF conversion
- [ ] Email template customization
- [ ] Archive management and search
- [ ] Modern .NET 6+ migration

---

**Status**: Work in Progress — no release package available yet  
**Repository**: [github.com/ofer-aha/SaveAsPDF](https://github.com/ofer-aha/SaveAsPDF)
