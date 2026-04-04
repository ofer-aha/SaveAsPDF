# Security Policy

## Project Status

> ⚠️ **WORK IN PROGRESS** — SaveAsPDF is under active development and has **no release package yet**. The security policy below reflects the current state of the source code.

## Supported Versions

As there are no official releases yet, security fixes are applied directly to the `master` branch. Once versioned releases are published, this table will be updated.

| Version | Supported |
|---------|-----------|
| master (pre-release) | ✅ Active development |
| Any tagged release | None yet |

## Security Features in the Current Codebase

The following security controls are already implemented and active in the source:

### Input Validation
- **Project ID validation** (`FileFoldersHelper.SafeProjectID`): All project IDs are validated against a strict regex before being used in any file-system path. Invalid IDs are rejected at both the UI and helper layers.
- **Folder/file name sanitization** (`FileFoldersHelper.SafeFolderName`): All folder and file names derived from user input or email metadata are scrubbed of invalid Windows file-system characters. Windows reserved device names (`CON`, `PRN`, `AUX`, `NUL`, `COM1`–`COM3`, `LPT1`–`LPT3`) are automatically renamed to avoid collisions.
- **Path normalization**: All file-system paths are resolved to their canonical form using `Path.GetFullPath` before use, preventing path traversal via relative components such as `..`.

### HTML Output (XSS Prevention)
- All user-supplied and email-derived data written into generated HTML files is encoded with `WebUtility.HtmlEncode` before insertion. This covers: project name, project ID, notes, user name, employee names, email addresses, and attachment file names.
- File paths embedded as `file:///` hyperlinks in generated HTML have `space`, `#`, `?`, and `%` percent-encoded to prevent link injection.

### File System Safety
- Unique file/folder name generation prevents silent overwrites when a target already exists.
- Attachment file names are sanitized (invalid characters stripped) before being written to disk.
- Password-protected Word documents are detected before conversion is attempted.
- Configuration metadata is stored in a **hidden directory** (`.SaveAsPDF`) to reduce accidental modification.
- Read-only/hidden file attributes on existing XML configuration files are cleared before overwriting to prevent silent write failures.

### Error Handling and Access Control
- `UnauthorizedAccessException` is explicitly caught on all file and directory operations and surfaced to the user without crashing the add-in.
- All Outlook COM objects (folders, items, stores) are released in `finally` blocks via `Marshal.ReleaseComObject`, preventing use-after-free and memory leaks.
- MAPI contact search filter strings escape single-quotes to prevent filter injection.
- Startup and selection-change exceptions are logged to `C:\Temp\SaveAsPDF_error.txt` for post-mortem debugging without exposing sensitive information in the UI.

### Data Storage
- All user data (project metadata, employee lists, preferences) is stored **locally** on the user's machine in XML files. No data is transmitted to any remote server.
- XML serialization uses `XmlReader`/`XmlWriter` with explicit UTF-8 encoding.

## Known Limitations and Areas for Future Improvement

The following are known gaps that are expected to be addressed before a stable release:

- The SMTP credentials in `MailLogic.cs` are placeholder strings (`"senderEmail"`, `"senderDisplayName"`) marked with `//TODO3`. Email sending is not functional and should not be used in production.
- Error log (`C:\Temp\SaveAsPDF_error.txt`) path is hardcoded. In a production release this should use a user-writable temp path and should not log internal stack traces to a world-readable location.
- No code-signing certificate is applied to the COM DLL yet. For production deployment, the assembly should be Authenticode-signed.

## Reporting a Vulnerability

If you discover a security vulnerability, please report it responsibly:

1. **Do not open a public GitHub issue** for security vulnerabilities.
2. Use one of the following private channels:
   - **GitHub Private Vulnerability Reporting**: [Security advisories](https://github.com/ofer-aha/SaveAsPDF/security/advisories/new)
   - **Discussions (private message)**: [Start a discussion](https://github.com/ofer-aha/SaveAsPDF/discussions)
   - **Email**: Contact via the GitHub profile

Please include:
- A description of the vulnerability and its potential impact
- Steps to reproduce or proof-of-concept code
- The version or commit SHA where the issue was found

We aim to acknowledge reports within **5 business days** and provide a fix or mitigation plan within **30 days** for confirmed issues.
