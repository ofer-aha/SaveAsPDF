# SaveAsPDF - Outlook Add-In

Convert Outlook mail messages to PDF files with ease. SaveAsPDF is a powerful Microsoft Outlook add-in that enables users to save emails and their attachments as PDF documents, with full support for project management, employee tracking, and advanced customization options.

## Features

- **📧 Email to PDF Conversion**: Convert any Outlook email message to a professional PDF file
- **📎 Attachment Management**: Select and save email attachments alongside the PDF
- **🏢 Project Management**: Organize emails by project with automatic project ID detection from email subjects
- **👥 Employee Directory**: Manage project team members with contact information and role assignment
- **📝 Rich Notes**: Add and maintain both mail-specific and project-specific notes
- **🌳 Folder Structure**: Browse and manage save locations with an intuitive folder tree view
- **🎨 Customization**: Configure fonts, colors, and application settings to your preference
- **🌍 Right-to-Left Support**: Full support for Hebrew and other RTL languages
- **🌙 Dark Mode Support**: Automatic theme adaptation to Windows system colors (light/dark modes)
- **⌨️ Quick Access**: Ribbon interface for quick conversion of selected messages

## System Requirements

- **Microsoft Outlook**: 2016 or later (Office 365 / Microsoft 365 compatible)
- **.NET Framework**: 4.7.2 or higher
- **Windows**: Windows 7 or later (32-bit or 64-bit)
- **RAM**: Minimum 512 MB (1 GB or more recommended)

## Installation

### Method 1: From Release Files
1. Download the latest release from the [Releases](https://github.com/ofer-aha/SaveAsPDF/releases) page
2. Run the installer
3. Follow the installation wizard
4. Restart Microsoft Outlook
5. The SaveAsPDF ribbon will appear in Outlook

### Method 2: From Source
1. Clone the repository:
   ```bash
   git clone https://github.com/ofer-aha/SaveAsPDF.git
   cd SaveAsPDF
   ```

2. Open `SaveAsPDF.sln` in Visual Studio 2019 or later

3. Restore NuGet packages:
   ```bash
   nuget restore SaveAsPDF.sln
   ```

4. Build the project (Release configuration recommended)

5. Deploy the add-in:
   - Visual Studio will automatically set up the add-in for development
   - For production deployment, use the ClickOnce installer in the Publish folder

## Quick Start

### Converting an Email to PDF

1. **Select an Email**: Open Outlook and click on the email you want to convert
2. **Click "Save as PDF"**: Look for the SaveAsPDF button in the Outlook ribbon (or use the Task Pane)
3. **Configure Options**:
   - Enter or select a project number (auto-detected from email subject when possible)
   - Choose attachments to include
   - Add notes or project-specific information
   - Select save location
4. **Save**: Click the "Save as PDF" button to generate the file

### Managing Projects

1. **Create a New Project**: Click "New Project" in the task pane
   - Set project number, name, and default save location
   - Add team member information

2. **Add Employees**: Click "Add Employee" or access the Employees tab
   - Manage team member details
   - Assign project leader role

3. **View Project Notes**: Switch between mail notes and project notes using the tabs
   - Mail-specific notes for individual emails
   - Project-wide notes for team coordination

## Architecture

### Project Structure

```
SaveAsPDF/
├── SaveAsPDF/                    # Main add-in project
│   ├── Helpers/                  # Helper classes for various operations
│   │   ├── MailLogic.cs
│   │   ├── SettingsHelpers.cs
│   │   ├── XmlFileHelper.cs
│   │   ├── TreeHelpers.cs
│   │   └── ...
│   ├── Models/                   # Data models
│   │   ├── ProjectModel.cs
│   │   ├── EmployeeModel.cs
│   │   ├── SettingsModel.cs
│   │   └── ...
│   ├── Forms/                    # User interface forms
│   │   ├── FormMain.cs           # Main dialog
│   │   ├── FormSettings.cs       # Settings dialog
│   │   ├── FormContacts.cs       # Employee management
│   │   └── FormNewProject.cs     # New project creation
│   ├── MainFormTaskPaneControl.cs # Task pane control (Outlook sidebar)
│   ├── ThisAddIn.cs              # Add-in entry point
│   ├── SaveAsPDFRibbon.cs        # Ribbon UI definition
│   └── Properties/               # Project properties and resources
├── SaveAsPDF.Tests/              # Unit tests
├── BenchmarkSuite1/              # Performance benchmarks
└── README.md                     # This file
```

### Technology Stack

- **Language**: C# 7.3
- **.NET Framework**: 4.7.2
- **Office Interop**: Microsoft Office Interop Libraries (Word, Outlook)
- **UI Framework**: Windows Forms
- **Configuration**: XML-based project settings
- **JSON Support**: Newtonsoft.Json for data serialization
- **Dialog Enhancement**: Ookii.Dialogs for native Windows dialogs

### Key Components

#### Core Classes

- **`MainFormTaskPaneControl`**: Task pane control for Outlook sidebar interface
  - Right-to-left layout support
  - Dynamic tab structure with nested tabs for notes
  - System color theme adaptation
  - Data grid views for attachments and employees

- **`MailLogic`**: Email processing and PDF conversion
  - Attachment extraction
  - HTML to PDF conversion
  - Email property extraction

- **`SettingsHelpers`**: Project and application settings management
  - XML-based configuration storage
  - Default paths and folders
  - User preferences

- **`XmlFileHelper`**: XML serialization and deserialization
  - Project model persistence
  - Employee list management
  - Settings storage

#### Data Models

- **`ProjectModel`**: Project information and settings
- **`EmployeeModel`**: Team member details with leadership assignment
- **`SettingsModel`**: Application configuration
- **`AttachmentsModel`**: Email attachment metadata
- **`DocumentModel`**: Converted document information

## Usage Examples

### Basic Email Conversion

```csharp
// Select email in Outlook and click "Save as PDF"
// The task pane will open with options to:
// - Select/auto-detect project
// - Choose attachments
// - Add notes
// - Pick save location
```

### Project Configuration

Email subjects with project IDs are automatically recognized. Example:
- Email subject: "Project P001234 - Meeting Notes"
- Auto-detected project ID: P001234

### Custom Settings

1. **Access Settings**: Click "Settings" in the task pane
2. **Configure**:
   - Default save location
   - Font preferences
   - Application behavior
   - PDF conversion options

## Configuration Files

SaveAsPDF uses XML-based configuration files stored in the project directories:

- **Project.xml**: Project information (name, notes, metadata)
- **Employees.xml**: Team member list and roles
- **Settings.xml**: User preferences and configurations
- **DefaultTree.fld**: Default folder structure template

## Interface Features

### Task Pane Layout

The Outlook sidebar provides:
- **Top Section**: Project and email metadata
  - Project ID and name
  - Email subject
  - Save location selector
  - File path display

- **Middle Section**: Tabbed content area
  - **Attachments**: Email file attachments with selection
  - **Employees**: Project team members directory
  - **Folders**: Hierarchical project folder browser
  - **Notes**: Sub-tabs for mail and project notes

- **Bottom Section**: Action buttons and status
  - Save to PDF, Cancel, Settings, New Project
  - Status messages and help text
  - PDF opening preference checkbox

### Theme Support

- **Light Mode**: Uses system `SystemColors.Window` and `SystemColors.Control`
- **Dark Mode**: Automatically adapts to Windows dark theme
- **RTL Layout**: Full right-to-left support for Hebrew and Arabic

## Development

### Prerequisites

- Visual Studio 2019 or later
- .NET Framework 4.7.2 Developer Pack
- Microsoft Office development tools
- NuGet package manager

### Building from Source

1. **Clone repository**:
   ```bash
   git clone https://github.com/ofer-aha/SaveAsPDF.git
   cd SaveAsPDF
   ```

2. **Restore packages**:
   ```bash
   nuget restore SaveAsPDF.sln
   ```

3. **Build solution**:
   ```bash
   msbuild SaveAsPDF.sln /p:Configuration=Release
   ```

4. **Run tests**:
   ```bash
   dotnet test SaveAsPDF.Tests.csproj
   ```

### Code Style

- **Language**: C# 7.3 features
- **Framework**: .NET Framework 4.7.2
- **Conventions**:
  - PascalCase for class and method names
  - camelCase for local variables and parameters
  - Prefix `_` for private fields
  - XML documentation comments for public members

### Testing

Unit tests are available in the `SaveAsPDF.Tests` project:

```bash
dotnet test SaveAsPDF.Tests.csproj
```

## Troubleshooting

### Add-in Not Appearing in Outlook

1. **Check if installed**: File → Options → Trust Center → Trust Center Settings → Trusted Add-ins
2. **Re-enable add-in**: Look for disabled add-ins in Trust Center
3. **Restart Outlook**: Close and reopen Microsoft Outlook
4. **Reinstall**: Uninstall and reinstall the add-in

### PDF Conversion Issues

- **Large attachments**: Some emails may take longer to process
- **Encoding issues**: Ensure email is in standard format (not RTF)
- **Path issues**: Verify write permissions to selected save location

### File Location Problems

- **UNC paths**: Network paths may require authentication
- **Special characters**: Avoid special characters in project names and file paths
- **Long paths**: Windows has a 260-character path limit (enable long paths in Windows 10+)

## Contributing

Contributions are welcome! Please follow these guidelines:

1. **Fork the repository**
2. **Create a feature branch**: `git checkout -b feature/YourFeature`
3. **Make your changes**: Follow the existing code style
4. **Write tests**: Add unit tests for new functionality
5. **Submit a Pull Request**: Describe your changes and reference any issues

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

For issues, questions, or suggestions:

- **GitHub Issues**: [Report a bug](https://github.com/ofer-aha/SaveAsPDF/issues)
- **Discussions**: [Start a discussion](https://github.com/ofer-aha/SaveAsPDF/discussions)
- **Email**: Contact via GitHub profile

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for version history and release notes.

## Credits

SaveAsPDF is maintained by the open-source community. Special thanks to:
- Contributors and bug reporters
- Microsoft Office Interop Libraries developers
- Third-party library maintainers

## Disclaimer

SaveAsPDF is provided as-is. The developers are not responsible for:
- Data loss or corruption
- PDF conversion quality or accuracy
- Compatibility issues with specific Outlook versions
- Third-party software conflicts

## Roadmap

Future planned features:
- [ ] Cloud storage integration (OneDrive, SharePoint)
- [ ] Email template customization
- [ ] Batch PDF conversion
- [ ] Archive management and search
- [ ] Enhanced reporting features
- [ ] Mobile companion app
- [ ] Modern .NET / .NET 6+ migration

---

**Version**: 1.0.0  
**Last Updated**: 2024  
**Repository**: [github.com/ofer-aha/SaveAsPDF](https://github.com/ofer-aha/SaveAsPDF)
