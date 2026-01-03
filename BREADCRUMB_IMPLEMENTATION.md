# Breadcrumb Address Bar Implementation

## Overview
The `txtFullPath` and `cmbSaveLocation` controls have been enhanced to display paths in a Windows Explorer-style breadcrumb format using arrow separators (›) instead of backslashes.

## Changes Made

### 1. Visual Styling

#### `txtFullPath` (Read-Only Display)
- Changed `BackColor` from `SystemColors.Control` to `SystemColors.Window` for address bar appearance
- Added `BorderStyle.FixedSingle` for a defined border
- Set `Cursor` to `Cursors.Arrow` to indicate read-only nature
- Displays paths in breadcrumb format: `C › Users › Admin › Documents`

#### `cmbSaveLocation` (Editable ComboBox)
- Kept as `ComboBoxStyle.DropDown` to allow both typing and selection
- Added Enter/Leave event handlers for smart breadcrumb display
- Stores actual path in `Tag` property while displaying breadcrumb in `Text`

### 2. New Helper Method

```csharp
private string FormatPathAsBreadcrumb(string path)
```
Converts standard Windows paths (`C:\Users\Admin`) into breadcrumb format (`C › Users › Admin`).

### 3. Smart Display Behavior

#### When User Enters `cmbSaveLocation`:
- Automatically converts breadcrumb display to full path for editing
- Example: `C › Users › Admin` becomes `C:\Users\Admin`

#### When User Leaves `cmbSaveLocation`:
- Converts typed path back to breadcrumb format
- Stores actual path in `Tag` property
- Example: User types `C:\Projects\New` → displays as `C › Projects › New`

#### When User Selects from Tree View:
- Path is automatically formatted as breadcrumb
- Both `cmbSaveLocation` and `txtFullPath` are updated

### 4. Path Retrieval in `btnOK_Click`

The click handler now:
1. First tries to get actual path from `cmbSaveLocation.Tag`
2. If Tag is null, uses `cmbSaveLocation.Text`
3. Converts any breadcrumb format back to standard path
4. Validates and processes the path as before

```csharp
string sPath = (cmbSaveLocation.Tag as string) ?? cmbSaveLocation.Text;
if (sPath.Contains(" › "))
{
    sPath = sPath.Replace(" › ", "\\");
}
```

## Benefits

1. **Modern Look**: Mimics Windows Explorer's address bar appearance
2. **Easy Reading**: Breadcrumb format is easier to scan visually
3. **Backward Compatible**: All existing functionality preserved
4. **No Breaking Changes**: Existing code that reads paths still works
5. **User-Friendly**: Automatically switches between breadcrumb display and editable path

## Technical Details

### Separator Character
- Uses › (U+203A Single Right-Pointing Angle Quotation Mark)
- This character is not a valid path character, making it safe for conversion

### Path Storage
- Actual path stored in `ComboBox.Tag` property
- Display text shows breadcrumb format
- Conversion happens automatically on Enter/Leave events

### Validation
- All existing path validation remains intact
- Breadcrumb format is converted before validation
- No changes to error handling logic

## Usage Examples

### Setting a Path Programmatically
```csharp
cmbSaveLocation.Tag = @"C:\Projects\SaveAsPDF";
cmbSaveLocation.Text = FormatPathAsBreadcrumb(@"C:\Projects\SaveAsPDF");
// Displays as: C › Projects › SaveAsPDF
```

### Getting the Actual Path
```csharp
string actualPath = (cmbSaveLocation.Tag as string) ?? cmbSaveLocation.Text;
if (actualPath.Contains(" › "))
{
    actualPath = actualPath.Replace(" › ", "\\");
}
```

## Testing Recommendations

1. Test path entry with various formats
2. Test selection from dropdown
3. Test tree view node clicking
4. Test folder picker dialog
5. Test with paths containing spaces
6. Test with Hebrew characters in path
7. Test undo/redo functionality
8. Verify PDF save location is correct

## Future Enhancements (Optional)

1. **Clickable Breadcrumb Segments**: Make each segment clickable to navigate
2. **Dropdown per Segment**: Show available folders when clicking a segment
3. **Copy/Paste Support**: Handle clipboard operations with both formats
4. **Drag-and-Drop**: Accept folder drops to set path
5. **Recent Paths**: Show recently used paths in a special format
