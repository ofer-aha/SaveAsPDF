// Ignore Spelling: Deserialize

using SaveAsPDF.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace SaveAsPDF.Helpers
{
    public static class XmlFileHelper
    {
        /// <summary>
        /// Convert the ProjectModel to XML file
        /// </summary>
        /// <param name="project">ProjectModel</param>
        /// <param name="path">The path for the .SaveAsPDF hidden folder</param>
        public static void ProjectModelToXmlFile(this string path, ProjectModel project)
        {
            // Create the XmlDocument.
            XmlDocument doc = new XmlDocument();
            doc.LoadXml("<Project></Project>");

            // Add a price element.
            XmlElement newElem = doc.CreateElement("ProjectNumber");
            newElem.InnerText = project.ProjectNumber;
            doc.DocumentElement.AppendChild(newElem);

            newElem = doc.CreateElement("ProjectName");
            newElem.InnerText = project.ProjectName;
            doc.DocumentElement.AppendChild(newElem);

            newElem = doc.CreateElement("NoteToProjectLeader");
            newElem.InnerText = project.NoteToProjectLeader.ToString();
            doc.DocumentElement.AppendChild(newElem);

            newElem = doc.CreateElement("ProjectNotes");
            newElem.InnerText = project.ProjectNotes;
            doc.DocumentElement.AppendChild(newElem);

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            // Save the document to a file and auto-indent the _employeesModel.


            FileFoldersHelper.CreateHiddenDirectory(path);

            XmlWriter writer = XmlWriter.Create(path, settings);

            try
            {
                doc.Save(writer);
                //close the writer
                writer.Close();

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "SaveAsPDF:ProjectModelToXmlFile");
            }
        }
        /// <summary>
        /// Converts Employee model to XML file
        /// </summary>
        /// <param name="employees">Employee Model</param>
        /// <param name="path">File Name as string</param>
        public static void EmployeesModelToXmlFile(this string path, List<EmployeeModel> employees)
        {
            XmlDocument xmlDocument = new XmlDocument();
            XmlNode docNode = xmlDocument.CreateXmlDeclaration("1.0", "UTF-8", null);
            xmlDocument.AppendChild(docNode);

            XmlNode employeesNode = xmlDocument.CreateElement("Employees");
            xmlDocument.AppendChild(employeesNode);

            foreach (EmployeeModel employee in employees)
            {
                XmlNode employeeNode = xmlDocument.CreateElement("Employee");
                employeesNode.AppendChild(employeeNode);

                XmlNode idNode = xmlDocument.CreateElement("ID");
                idNode.AppendChild(xmlDocument.CreateTextNode(employee.Id.ToString()));
                employeeNode.AppendChild(idNode);

                XmlNode firstNameNode = xmlDocument.CreateElement("FirstName");
                firstNameNode.AppendChild(xmlDocument.CreateTextNode(employee.FirstName));
                employeeNode.AppendChild(firstNameNode);

                XmlNode lastNameNode = xmlDocument.CreateElement("LastName");
                lastNameNode.AppendChild(xmlDocument.CreateTextNode(employee.LastName));
                employeeNode.AppendChild(lastNameNode);

                XmlNode emailAddressNode = xmlDocument.CreateElement("EmailAddress");
                emailAddressNode.AppendChild(xmlDocument.CreateTextNode(employee.EmailAddress));
                employeeNode.AppendChild(emailAddressNode);
            }
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            // Save the document to a file and auto-indent the _employeesModel.
            XmlWriter writer = XmlWriter.Create(path, settings);
            try
            {
                xmlDocument.Save(writer);
                //close the writer
                writer.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SaveAsPDF:EmployeesModelToXmlFile");
            }
        }

        /// <summary>
        /// Extension method Converts the Project XML file to ProjectModel
        /// </summary>
        /// <param name="xmlFile">Source file name</param>
        /// <returns><see cref="ProjectModel"/></returns>
        public static ProjectModel XmlProjectFileToModel(this string xmlFile)
        {
            ProjectModel projectModel = new ProjectModel();
            XmlDocument xmlDoc = new XmlDocument();

            try
            {
                xmlDoc.Load(xmlFile);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "SaveAsPDF:XmlProjectFileToModel");
                return null;
            }
            XmlNodeList projectName = xmlDoc.GetElementsByTagName("ProjectName");
            XmlNodeList NoteToProjectLeader = xmlDoc.GetElementsByTagName("NoteToProjectLeader");
            XmlNodeList projectNotes = xmlDoc.GetElementsByTagName("ProjectNotes");

            projectModel.ProjectName = projectName[0].InnerText;
            projectModel.NoteToProjectLeader = bool.Parse(NoteToProjectLeader[0].InnerText);
            projectModel.ProjectNotes = projectNotes[0].InnerText;


            return projectModel;
        }

        /// <summary>
        /// Extension method Import the XML file to employee model
        /// </summary>
        /// <param name="xmlFile">Full path to XML file</param>
        /// <returns><see cref="ProjectModel"/></returns>
        public static List<EmployeeModel> XmlEmployeesFileToModel(this string xmlFile)
        {
            List<EmployeeModel> employeesModel = new List<EmployeeModel>();
            XmlDocument xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.Load(xmlFile);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "SaveAsPDF:XmlEmployeesFileToModel");
                return null;
            }

            XmlNodeList FirstName = xmlDoc.GetElementsByTagName("FirstName");
            XmlNodeList LastName = xmlDoc.GetElementsByTagName("LastName");
            XmlNodeList EmailAddress = xmlDoc.GetElementsByTagName("EmailAddress");

            for (int i = 0; i < FirstName.Count; i++)
            {
                EmployeeModel em = new EmployeeModel
                {
                    FirstName = FirstName[i].InnerXml.ToString(),
                    LastName = LastName[i].InnerXml.ToString(),
                    EmailAddress = EmailAddress[i].InnerXml.ToString()
                };
                employeesModel.Add(em);
            }

            //xmlDoc.Save(xmlFile);
            return employeesModel;
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
