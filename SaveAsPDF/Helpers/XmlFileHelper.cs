// Ignore Spelling: Deserialize

using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

/// <summary>
/// Provides helper methods for serializing and deserializing project and employee models to and from XML files.
/// </summary>
namespace SaveAsPDF.Helpers
{
    public static class XmlFileHelper
    {
        /// <summary>
        /// Serializes a <see cref="ProjectModel"/> to an XML file at the specified file path.
        /// </summary>
        /// <param name="filePath">The full path to the XML file where the <see cref="ProjectModel"/> will be saved.</param>
        /// <param name="projectModel">The <see cref="ProjectModel"/> object to serialize and save.</param>
        public static void ProjectModelToXmlFile(this string filePath, ProjectModel projectModel)
        {
            try
            {
                string directoryPath = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directoryPath) && !Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                XmlSerializer serializer = new XmlSerializer(typeof(ProjectModel));
                using (StreamWriter writer = new StreamWriter(filePath, false))
                {
                    serializer.Serialize(writer, projectModel);
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
            try
            {
                FileFoldersHelper.CreateHiddenDirectory(Path.GetDirectoryName(path));

                XmlSerializer serializer = new XmlSerializer(typeof(List<EmployeeModel>));
                using (StreamWriter writer = new StreamWriter(path))
                {
                    serializer.Serialize(writer, employees);
                }
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"אירעה שגיאה בשמירת רשימת העובדים: {ex.Message}",
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
        /// <param name="xmlFile">The full path to the XML file to load.</param>
        /// <returns>
        /// The deserialized <see cref="ProjectModel"/> if successful; otherwise, <c>null</c> if an error occurs.
        /// </returns>
        public static ProjectModel XmlProjectFileToModel(this string xmlFile)
        {
            try
            {
                if (!File.Exists(xmlFile))
                {
                    throw new FileNotFoundException($"הקובץ '{xmlFile}' לא קיים.");
                }

                var xmlDoc = new System.Xml.XmlDocument();
                xmlDoc.Load(xmlFile);

                var project = new ProjectModel
                {
                    ProjectNumber = xmlDoc.SelectSingleNode("//ProjectModel/ProjectNumber")?.InnerText,
                    ProjectName = xmlDoc.SelectSingleNode("//ProjectModel/ProjectName")?.InnerText,
                    NoteToProjectLeader = bool.Parse(xmlDoc.SelectSingleNode("//ProjectModel/NoteToProjectLeader")?.InnerText ?? "false"),
                    ProjectNotes = xmlDoc.SelectSingleNode("//ProjectModel/ProjectNotes")?.InnerText
                };

                return project;
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"אירעה שגיאה בטעינת הפרויקט: {ex.Message}",
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
        /// <param name="xmlFile">The full path to the XML file to load.</param>
        /// <returns>
        /// The deserialized list of <see cref="EmployeeModel"/> objects if successful; otherwise, <c>null</c> if an error occurs.
        /// </returns>
        public static List<EmployeeModel> XmlEmployeesFileToModel(this string xmlFile)
        {
            try
            {
                if (!File.Exists(xmlFile))
                {
                    throw new FileNotFoundException($"הקובץ '{xmlFile}' לא קיים.");
                }

                XmlSerializer serializer = new XmlSerializer(typeof(List<EmployeeModel>));
                using (FileStream fileStream = new FileStream(xmlFile, FileMode.Open))
                {
                    return (List<EmployeeModel>)serializer.Deserialize(fileStream);
                }
            }
            catch (InvalidOperationException ex)
            {
                XMessageBox.Show(
                    $"אירעה שגיאה בעת המרת קובץ ה-XML: {ex.InnerException?.Message ?? ex.Message}",
                    "SaveAsPDF:XmlEmployeesFileToModel",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return null;
            }
            catch (Exception ex)
            {
                XMessageBox.Show(
                    $"אירעה שגיאה בטעינת העובדים: {ex.Message}",
                    "SaveAsPDF:XmlEmployeesFileToModel",
                    XMessageBoxButtons.OK,
                    XMessageBoxIcon.Error,
                    XMessageAlignment.Right,
                    XMessageLanguage.Hebrew
                );
                return null;
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
