using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace SaveAsPDF.Helpers
{
    public static class ComboBoxExtensions
    {
        // Cache for directory existence checks to improve performance in CustomizeComboBox
        private static readonly Dictionary<string, CacheEntry> DirectoryExistsCache = new Dictionary<string, CacheEntry>(StringComparer.OrdinalIgnoreCase);
        private const int MaxCacheSize = 1000; // Maximum number of entries to store
        private static readonly TimeSpan CacheExpiration = TimeSpan.FromMinutes(5); // Cache expires after 5 minutes

        // Simple struct to store cached directory information with timestamp
        private struct CacheEntry
        {
            public bool Exists;
            public DateTime Timestamp;

            public CacheEntry(bool exists)
            {
                Exists = exists;
                Timestamp = DateTime.Now;
            }

            public bool IsExpired => DateTime.Now - Timestamp > CacheExpiration;
        }

        /// <summary>
        /// Loads a ComboBox with paths from a specified file, replacing placeholders with the ProjectID and current date.
        /// </summary>
        /// <param name="comboBox">The combo box to load the paths into.</param>
        /// <param name="fileName">The name of the file containing the paths.</param>
        /// <param name="projectID">The ProjectID to filter or modify the paths.</param>
        /// <param name="settings">The settings model containing configuration values. If null, uses FormMain.settingsModel.</param>
        public static void LoadComboBoxWithPaths(this ComboBox comboBox, string fileName, string projectID, Models.SettingsModel settings = null)
        {
            if (comboBox == null)
            {
                XMessageBox.Show(
                    "אובייקט ה-ComboBox לא יכול להיות ריק.",
                    "שגיאת קלט",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return;
            }

            if (string.IsNullOrWhiteSpace(projectID))
            {
                XMessageBox.Show(
                    "מזהה הפרויקט לא יכול להיות ריק.",
                    "שגיאת קלט",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return;
            }

            // Use provided settings or fall back to global settings
            var settingsModel = settings ?? FormMain.settingsModel;
            
            // Create project folder path based on Project ID - format is RootDrive\xx\ProjectID\
            // Where xx is the first two characters of the project ID or with a leading 0 for 3-character IDs
            string projectFolder;
            
            try
            {
                // Get the project root folder path - this should already be in the correct format (RootDrive\xx\ProjectID\)
                // from the ProjectFullPath method in FileFoldersHelper.cs
                projectFolder = settingsModel.ProjectRootFolder?.FullName;
                
                // If it's null or doesn't exist, manually construct it (as a fallback)
                if (string.IsNullOrEmpty(projectFolder))
                {
                    string rootPath = settingsModel.RootDrive.TrimEnd('\\') + "\\";
                    string prefix;
                    
                    // Use correct two-digit folder prefix based on project ID format
                    if (projectID.Length == 3)
                        prefix = $"0{projectID.Substring(0, 1)}"; // e.g., 123 -> 01
                    else
                        prefix = projectID.Substring(0, 2); // e.g., 1234 -> 12
                    
                    projectFolder = $"{rootPath}{prefix}\\{projectID}\\";
                }
            }
            catch (Exception)
            {
                // If any error occurs, use a simple fallback path
                projectFolder = Path.Combine(settingsModel.RootDrive, projectID);
            }
            
            comboBox.BeginUpdate();
            comboBox.Items.Clear();
            
            try
            {
                // First add the project root folder itself
                if (!string.IsNullOrEmpty(projectFolder) && !comboBox.Items.Contains(projectFolder))
                {
                    comboBox.Items.Add(projectFolder);
                }
                
                // Build standard paths for the project
                string[] standardPaths = new string[]
                {
                    Path.Combine(projectFolder, "DWG"),
                    Path.Combine(projectFolder, "PDF"),
                    Path.Combine(projectFolder, "OLD"),
                    Path.Combine(projectFolder, "תכניות")
                };
                
                // Use HashSet for faster contains checks
                HashSet<string> addedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                addedPaths.Add(projectFolder);
                
                foreach (string path in standardPaths)
                {
                    if (!addedPaths.Contains(path))
                    {
                        comboBox.Items.Add(path);
                        addedPaths.Add(path);
                    }
                }
                
                // If a file is specified and it exists, load additional paths
                if (!string.IsNullOrEmpty(fileName) && File.Exists(fileName))
                {
                    try
                    {
                        string[] lines = File.ReadAllLines(fileName);
                        string currentDate = DateTime.Now.ToString("dd.MM.yyyy");
                        
                        foreach (string line in lines)
                        {
                            if (string.IsNullOrWhiteSpace(line)) continue;
                            
                            try
                            {
                                // Replace tags with actual values
                                string path = line
                                    .Replace(settingsModel.ProjectRootTag, projectID)
                                    .Replace(settingsModel.DateTag, currentDate);
                                    
                                // Handle path construction based on whether it's relative or absolute
                                if (Path.IsPathRooted(path))
                                {
                                    // For absolute paths, check if they already contain the project folder structure
                                    // to avoid duplicating the project ID in the path
                                    
                                    // First check if projectID appears in the path already
                                    string projectCheck = $"\\{projectID}\\";
                                    if (path.Contains(projectCheck))
                                    {
                                        // The path already contains the project ID correctly, use as-is
                                        // Clean up any double backslashes
                                        path = path.Replace("\\\\", "\\").TrimEnd('\\') + "\\";
                                    }
                                    else
                                    {
                                        // Path doesn't have project ID yet, so just ensure proper formatting
                                        path = path.Replace("\\\\", "\\").TrimEnd('\\') + "\\";
                                    }
                                }
                                else
                                {
                                    // For relative paths, combine with project folder but prevent duplicating the project ID
                                    if (path.StartsWith("\\") || path.StartsWith("/"))
                                        path = path.Substring(1);
                                    
                                    // Check if path already starts with project ID to avoid duplication
                                    if (path.StartsWith(projectID + "\\") || path.StartsWith(projectID + "/"))
                                    {
                                        // Path already has project ID, so we need to get parent folder of projectFolder
                                        string projectParentFolder = Path.GetDirectoryName(projectFolder.TrimEnd('\\'));
                                        path = Path.Combine(projectParentFolder, path);
                                    }
                                    else
                                    {
                                        // Normal relative path, just combine with project folder
                                        path = Path.Combine(projectFolder, path);
                                    }
                                    
                                    path = path.Replace("\\\\", "\\").TrimEnd('\\') + "\\";
                                }
                                
                                // Final check: Verify we don't have duplicate project IDs in the path
                                // e.g. "C:\10\1000\1000\..." -> should be "C:\10\1000\..."
                                string folderStructure = $"\\{projectID}\\{projectID}\\";
                                if (path.Contains(folderStructure))
                                {
                                    path = path.Replace(folderStructure, $"\\{projectID}\\");
                                }
                                
                                // Add the path if it's valid and not already in the list
                                if (IsValidPath(path) && !addedPaths.Contains(path))
                                {
                                    comboBox.Items.Add(path);
                                    addedPaths.Add(path);
                                }
                            }
                            catch
                            {
                                // Skip paths that cause exceptions
                                continue;
                            }
                        }
                    }
                    catch
                    {
                        // Skip file reading errors
                    }
                }
                
                // Set the default selection based on the DefaultSavePath setting
                if (comboBox.Items.Count > 0)
                {
                    // First try to find exact match for DefaultSavePath
                    string defaultPath = settingsModel.DefaultSavePath;
                    int index = -1;
                    
                    // Look for the DefaultSavePath in the combobox items
                    if (!string.IsNullOrEmpty(defaultPath))
                    {
                        // Check for exact match first
                        for (int i = 0; i < comboBox.Items.Count; i++)
                        {
                            if (string.Equals(comboBox.Items[i].ToString(), defaultPath, StringComparison.OrdinalIgnoreCase))
                            {
                                index = i;
                                break;
                            }
                        }
                        
                        // If no exact match, check if any path ends with the default path
                        // This handles cases where DefaultSavePath might be a relative path
                        if (index < 0)
                        {
                            for (int i = 0; i < comboBox.Items.Count; i++)
                            {
                                string itemPath = comboBox.Items[i].ToString();
                                if (itemPath.EndsWith(defaultPath.TrimEnd('\\') + "\\", StringComparison.OrdinalIgnoreCase))
                                {
                                    index = i;
                                    break;
                                }
                            }
                        }
                    }
                    
                    // If DefaultSavePath wasn't found or is empty, use first item
                    if (index < 0)
                    {
                        // Default to project folder if nothing else
                        comboBox.SelectedIndex = 0;
                    }
                    else
                    {
                        comboBox.SelectedIndex = index;
                    }
                }
            }
            finally
            {
                comboBox.EndUpdate();
            }
        }

        /// <summary>
        /// Validates whether a given path is in a legal format.
        /// </summary>
        /// <param name="path">The path to validate.</param>
        /// <returns>True if the path is valid; otherwise, false.</returns>
        private static bool IsValidPath(string path)
        {
            // Quick check for null or empty paths
            if (string.IsNullOrWhiteSpace(path))
                return false;

            // Check for invalid characters directly without try-catch for better performance
            if (path.IndexOfAny(Path.GetInvalidPathChars()) != -1)
                return false;

            // Additional validation for path format
            try
            {
                // Check if the path is rooted (absolute path)
                return Path.IsPathRooted(path);
            }
            catch
            {
                // Fall back to exception handling only when necessary
                return false;
            }
        }

        /// <summary>
        /// Customizes the ComboBox to draw items with different colors based on their content.
        /// </summary>
        /// <param name="comboBox">The ComboBox to customize.</param>
        public static void CustomizeComboBox(this ComboBox comboBox)
        {
            if (comboBox == null)
            {
                XMessageBox.Show(
                    "אובייקט ה-ComboBox לא יכול להיות ריק.",
                    "שגיאת קלט",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return;
            }

            // Reset event handler to prevent duplicates
            comboBox.DrawMode = DrawMode.OwnerDrawFixed;

            // Remove existing handlers before adding a new one
            comboBox.DrawItem -= ComboBox_DrawItem;

            // Add the handler
            comboBox.DrawItem += ComboBox_DrawItem;
        }

        // Extracted event handler to avoid creating multiple anonymous delegates
        private static void ComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            // Check if the index is valid
            if (e.Index < 0) return;

            var comboBox = sender as ComboBox;
            if (comboBox == null) return;

            // Get the item text
            string itemText = comboBox.Items[e.Index].ToString();

            // Determine the color based on directory existence with caching for performance
            Color textColor = !DirectoryExists(itemText) ? Color.Red : Color.Black;

            // Draw the background
            e.DrawBackground();

            // Draw the text with the specified color
            using (Brush brush = new SolidBrush(textColor))
            {
                e.Graphics.DrawString(itemText, e.Font, brush, e.Bounds);
            }

            // Draw the focus rectangle if the item is selected
            e.DrawFocusRectangle();
        }

        // Helper method to cache directory existence checks with size limits and expiration
        private static bool DirectoryExists(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;

            // Ensure cache doesn't grow too large
            if (DirectoryExistsCache.Count > MaxCacheSize)
            {
                ClearDirectoryCache();
            }

            // Use the cache for repeated checks
            if (DirectoryExistsCache.TryGetValue(path, out CacheEntry entry))
            {
                // Check if the entry has expired
                if (entry.IsExpired)
                {
                    // Update the cache with fresh information
                    bool exists = Directory.Exists(path);
                    DirectoryExistsCache[path] = new CacheEntry(exists);
                    return exists;
                }
                return entry.Exists;
            }

            // Check if it's a valid directory path
            bool directoryExists = Directory.Exists(path);

            // Cache the result
            DirectoryExistsCache[path] = new CacheEntry(directoryExists);

            return directoryExists;
        }

        /// <summary>
        /// Clears the directory existence cache or removes expired entries.
        /// </summary>
        /// <param name="removeExpiredOnly">If true, only removes expired entries; otherwise, clears the entire cache.</param>
        public static void ClearDirectoryCache(bool removeExpiredOnly = false)
        {
            if (removeExpiredOnly)
            {
                List<string> keysToRemove = new List<string>();
                foreach (var kvp in DirectoryExistsCache)
                {
                    if (kvp.Value.IsExpired)
                    {
                        keysToRemove.Add(kvp.Key);
                    }
                }

                foreach (string key in keysToRemove)
                {
                    DirectoryExistsCache.Remove(key);
                }
            }
            else
            {
                DirectoryExistsCache.Clear();
            }
        }

        /// <summary>
        /// Sets the ComboBox's text to a breadcrumb representation of a Windows Explorer path
        /// and populates its items with each folder segment.
        /// </summary>
        /// <param name="comboBox">The target ComboBox.</param>
        /// <param name="path">Optional explicit path to use. If null, uses the ComboBox's SelectedItem.</param>
        public static void SetBreadcrumbPath(this ComboBox comboBox, string path = null)
        {
            if (comboBox == null)
            {
                XMessageBox.Show(
                    "אובייקט ה-ComboBox לא יכול להיות ריק.",
                    "שגיאת קלט",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return;
            }

            // Get path from parameter or SelectedItem
            string pathToProcess = path;

            // If no explicit path provided, use SelectedItem
            if (pathToProcess == null)
            {
                if (comboBox.SelectedItem == null)
                    return;

                pathToProcess = comboBox.SelectedItem as string;
                if (string.IsNullOrWhiteSpace(pathToProcess))
                    return;
            }

            comboBox.BeginUpdate();
            try
            {
                // Trim trailing backslashes to avoid empty segments
                string trimmedPath = pathToProcess.TrimEnd('\\');

                // Split the path into segments using the backslash as the delimiter
                string[] segments = trimmedPath.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);

                // Create a breadcrumb string with appropriate initial capacity
                var breadcrumb = new System.Text.StringBuilder(trimmedPath.Length + segments.Length * 3);
                for (int i = 0; i < segments.Length; i++)
                {
                    if (i > 0) breadcrumb.Append(" > ");
                    breadcrumb.Append(segments[i]);
                }

                // Set the ComboBox's text to the breadcrumb representation
                comboBox.Text = breadcrumb.ToString();

                // Clear and populate items
                comboBox.Items.Clear();
                comboBox.Items.AddRange(segments);
            }
            finally
            {
                comboBox.EndUpdate();
            }
        }
    }
}
