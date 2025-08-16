
using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

/// <summary>
/// Provides helper methods for serializing and deserializing project and employee models to and from XML files.
/// </summary>
namespace SaveAsPDF.Helpers
{
    public static class XmlFileHelper
    {
        // Cache serializers for better performance
        private static readonly XmlSerializer ProjectSerializer = new XmlSerializer(typeof(ProjectModel));
        private static readonly XmlSerializer EmployeeListSerializer = new XmlSerializer(typeof(List<EmployeeModel>));
        
        // XML serializer settings
        private static readonly XmlWriterSettings WriterSettings = new XmlWriterSettings
        {
            Indent = true,
            IndentChars = "  ",
            NewLineChars = "\r\n",
            NewLineHandling = NewLineHandling.Replace,
            OmitXmlDeclaration = false,
            Encoding = System.Text.Encoding.UTF8
        };
        
        /// <summary>
        /// Serializes a <see cref="ProjectModel"/> to an XML file at the specified file path.
        /// </summary>
        /// <param name="filePath">The full path to the XML file where the <see cref="ProjectModel"/> will be saved.</param>
        /// <param name="projectModel">The <see cref="ProjectModel"/> object to serialize and save.</param>
        public static void ProjectModelToXmlFile(this string filePath, ProjectModel projectModel)
        {
            if (projectModel == null)
            {
                XMessageBox.Show(
                    "מודל הפרויקט אינו יכול להיות ריק.",
                    "SaveAsPDF:ProjectModelToXmlFile",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return;
            }
            
            try
            {
                // Create directory if it doesn't exist
                string directoryPath = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directoryPath) && !Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                // Use XmlWriter with optimal settings for better performance
                using (XmlWriter writer = XmlWriter.Create(filePath, WriterSettings))
                {
                    ProjectSerializer.Serialize(writer, projectModel);
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                XMessageBox.Show(
                    $"הגישה לנתיב נדחתה: {ex.Message}",
                    "SaveAsPDF:ProjectModelToXmlFile",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"אירעה שגיאה בשמירת מודל הפרויקט ל-XML: {ex.Message}",
                    "SaveAsPDF:ProjectModelToXmlFile",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }

        /// <summary>
        /// Serializes a list of <see cref="EmployeeModel"/> objects to an XML file at the specified path.
        /// </summary>
        /// <param name="path">The full path to the XML file where the employee list will be saved.</param>
        /// <param name="employees">The list of <see cref="EmployeeModel"/> objects to serialize and save.</param>
        public static void EmployeesModelToXmlFile(this string path, List<EmployeeModel> employees)
        {
            if (employees == null)
            {
                XMessageBox.Show(
                    "רשימת העובדים אינה יכולה להיות ריקה.",
                    "SaveAsPDF:EmployeesModelToXmlFile",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return;
            }
            
            try
            {
                FileFoldersHelper.CreateHiddenDirectory(Path.GetDirectoryName(path));
                
                // Use XmlWriter with optimal settings
                using (XmlWriter writer = XmlWriter.Create(path, WriterSettings))
                {
                    EmployeeListSerializer.Serialize(writer, employees);
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                XMessageBox.Show(
                    $"הגישה לנתיב נדחתה: {ex.Message}",
                    "SaveAsPDF:EmployeesModelToXmlFile",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"אירעה שגיאה בשמירת רשימת העובדים ל-XML: {ex.Message}",
                    "SaveAsPDF:EmployeesModelToXmlFile",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
            }
        }

        /// <summary>
        /// Deserializes a <see cref="ProjectModel"/> from an XML file at the specified path.
        /// </summary>
        /// <param name="filePath">The full path to the XML file to deserialize.</param>
        /// <returns>The deserialized <see cref="ProjectModel"/> object, or null if deserialization fails.</returns>
        public static ProjectModel XmlProjectFileToModel(string filePath)
        {
            if (!File.Exists(filePath))
            {
                XMessageBox.Show(
                    $"הקובץ {filePath} אינו קיים.",
                    "SaveAsPDF:XmlProjectFileToModel",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return null;
            }

            try
            {
                // Use XmlReader for more efficient reading
                using (XmlReader reader = XmlReader.Create(filePath))
                {
                    return (ProjectModel)ProjectSerializer.Deserialize(reader);
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"שגיאה בטעינת מודל הפרויקט מה-XML: {ex.Message}",
                    "SaveAsPDF:XmlProjectFileToModel",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return null;
            }
        }

        /// <summary>
        /// Deserializes a list of <see cref="EmployeeModel"/> objects from an XML file at the specified path.
        /// </summary>
        /// <param name="filePath">The full path to the XML file to deserialize.</param>
        /// <returns>The deserialized list of <see cref="EmployeeModel"/> objects, or an empty list if deserialization fails.</returns>
        public static List<EmployeeModel> XmlEmployeesFileToModel(this string filePath)
        {
            if (!File.Exists(filePath))
            {
                return new List<EmployeeModel>();
            }

            try
            {
                // Use XmlReader for more efficient reading
                using (XmlReader reader = XmlReader.Create(filePath))
                {
                    return (List<EmployeeModel>)EmployeeListSerializer.Deserialize(reader);
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"שגיאה בטעינת רשימת העובדים מה-XML: {ex.Message}",
                    "SaveAsPDF:XmlEmployeesFileToModel",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return new List<EmployeeModel>();
            }
        }

        /// <summary>
        /// Serializes a list of <see cref="SettingsModel"/> objects to an XML string.
        /// </summary>
        /// <param name="list">The list of <see cref="SettingsModel"/> objects to serialize.</param>
        /// <returns>
        /// A string containing the XML representation of the list.
        /// </returns>
        public static string SerializeList(List<SettingsModel> list)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(List<SettingsModel>));
            using (StringWriter writer = new StringWriter())
            {
                serializer.Serialize(writer, list);
                return writer.ToString();
            }
        }

        /// <summary>
        /// Deserializes a list of <see cref="SettingsModel"/> objects from an XML string.
        /// </summary>
        /// <param name="serializedList">The XML string representing the list of <see cref="SettingsModel"/> objects.</param>
        /// <returns>
        /// The deserialized list of <see cref="SettingsModel"/> objects.
        /// </returns>
        public static List<SettingsModel> DeserializeList(string serializedList)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(List<SettingsModel>));
            using (StringReader reader = new StringReader(serializedList))
            {
                return (List<SettingsModel>)serializer.Deserialize(reader);
            }
        }
    }
}
