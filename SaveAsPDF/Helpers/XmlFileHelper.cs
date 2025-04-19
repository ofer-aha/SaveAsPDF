// Ignore Spelling: Deserialize

using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace SaveAsPDF.Helpers
{
    public static class XmlFileHelper
    {
        /// <summary>
        /// Converts the ProjectModel to an XML file and saves it to the specified file path.
        /// </summary>
        /// <param name="filePath">The full path to the XML file where the ProjectModel will be saved.</param>
        /// <param name="projectModel">The ProjectModel object to be serialized and saved.</param>
        public static void ProjectModelToXmlFile(this string filePath, ProjectModel projectModel)
        {
            try
            {
                // Ensure the directory exists
                string directoryPath = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directoryPath) && !Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                // Serialize the ProjectModel to XML and save it to the file
                XmlSerializer serializer = new XmlSerializer(typeof(ProjectModel));
                using (StreamWriter writer = new StreamWriter(filePath))
                {
                    serializer.Serialize(writer, projectModel);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while saving the project model to XML: {ex.Message}",
                                "SaveAsPDF:ProjectModelToXmlFile",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Converts Employee model to XML file
        /// </summary>
        /// <param name="employees">Employee Model</param>
        /// <param name="path">File Name as string</param>
        public static void EmployeesModelToXmlFile(this string path, List<EmployeeModel> employees)
        {
            try
            {
                // Ensure the directory exists
                FileFoldersHelper.CreateHiddenDirectory(Path.GetDirectoryName(path));

                // Serialize the EmployeeModel list to XML
                XmlSerializer serializer = new XmlSerializer(typeof(List<EmployeeModel>));
                using (StreamWriter writer = new StreamWriter(path))
                {
                    serializer.Serialize(writer, employees);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SaveAsPDF:EmployeesModelToXmlFile");
            }
        }

        /// <summary>
        /// Extension method to import the XML file into a ProjectModel.
        /// </summary>
        /// <param name="xmlFile">Full path to the XML file.</param>
        /// <returns><see cref="ProjectModel"/> containing the project data.</returns>
        public static ProjectModel XmlProjectFileToModel(this string xmlFile)
        {
            try
            {
                // Check if the file exists
                if (!File.Exists(xmlFile))
                {
                    throw new FileNotFoundException($"The file '{xmlFile}' does not exist.");
                }

                // Load the XML file and parse it into a ProjectModel
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
                MessageBox.Show($"An error occurred while loading the project: {ex.Message}", "SaveAsPDF:XmlProjectFileToModel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }







        /// <summary>
        /// Extension method to import the XML file into a list of EmployeeModel.
        /// </summary>
        /// <param name="xmlFile">Full path to the XML file.</param>
        /// <returns><see cref="List{EmployeeModel}"/> containing the employees.</returns>

        public static List<EmployeeModel> XmlEmployeesFileToModel(this string xmlFile)
        {
            try
            {
                // Check if the file exists
                if (!File.Exists(xmlFile))
                {
                    throw new FileNotFoundException($"The file '{xmlFile}' does not exist.");
                }

                // Deserialize the XML file into a list of EmployeeModel
                XmlSerializer serializer = new XmlSerializer(typeof(List<EmployeeModel>), new XmlRootAttribute("ArrayOfEmployeeModel"));
                using (FileStream fileStream = new FileStream(xmlFile, FileMode.Open))
                {
                    return (List<EmployeeModel>)serializer.Deserialize(fileStream);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while loading employees: {ex.Message}", "SaveAsPDF:XmlEmployeesFileToModel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }


        /// <summary>
        /// Serialize the list to a string 
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
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
        /// De-serialize the string back to a list
        /// </summary>
        /// <param name="serializedList"></param>
        /// <returns><see cref="SettingsModel"/> Object</returns>
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
