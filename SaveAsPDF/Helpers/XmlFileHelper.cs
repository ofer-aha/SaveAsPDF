using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using System.Xml.Serialization;
using SaveAsPDF.Models;
using System.IO;
using System.Xml.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace SaveAsPDF.Helpers
{
   public static class XmlFileHelper
    {
       /// <summary>
       /// Convert the ProjectModel to XML file
       /// </summary>
       /// <param name="project">ProjectModel</param>
       /// <param name="path">The path for the .SaveAsPDF hidden folder</param>
        public static void ProjectModelToXmlFile(this string path ,ProjectModel project)
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

            newElem = doc.CreateElement("NoteEmployee");
            newElem.InnerText = project.NoteEmployee.ToString();
            doc.DocumentElement.AppendChild(newElem);

            newElem = doc.CreateElement("ProjectNotes");
            newElem.InnerText = project.ProjectNotes;
            doc.DocumentElement.AppendChild(newElem);

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            // Save the document to a file and auto-indent the output.
            XmlWriter writer = XmlWriter.Create(path, settings);
            
            try
            {
                doc.Save(writer);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "SaveAsPDF:ProjectModelToXmlFile");
            }
        }
        /// <summary>
        /// Converts Emploee model to XML file
        /// </summary>
        /// <param name="employees">Employee Model</param>
        /// <param name="path">File Name as string</param>
        public static void EmployeesModelToXmlFile(this string path, List<EmployeeModel> employees)
        {

            XmlDocument doc = new XmlDocument();
            XmlNode docNode = doc.CreateXmlDeclaration("1.0","UTF-8",null);
            doc.AppendChild(docNode);

            XmlNode employeesNode = doc.CreateElement("Employees");
            doc.AppendChild(employeesNode);

            foreach (EmployeeModel employee in employees)
            {

                XmlNode employeeNode = doc.CreateElement("Employee");
                employeesNode.AppendChild(employeeNode);

                XmlNode idNode = doc.CreateElement("ID");
                idNode.AppendChild(doc.CreateTextNode(employee.Id.ToString()));
                employeeNode.AppendChild(idNode);
                
                XmlNode firstNameNode = doc.CreateElement("FirstName");
                firstNameNode.AppendChild(doc.CreateTextNode(employee.FirstName));
                employeeNode.AppendChild(firstNameNode);

                XmlNode lastNameNode = doc.CreateElement("LastName");
                lastNameNode.AppendChild(doc.CreateTextNode(employee.LastName));
                employeeNode.AppendChild(lastNameNode);

                XmlNode emailAddressNode = doc.CreateElement("EmailAddress");
                emailAddressNode.AppendChild(doc.CreateTextNode(employee.EmailAddress));
                employeeNode.AppendChild(emailAddressNode);
               
                
            }
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            // Save the document to a file and auto-indent the output.
            XmlWriter writer = XmlWriter.Create(path, settings);
            
            try
            {
                doc.Save(writer);
            }
            
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "SaveAsPDF:EmployeesModelToXmlFile");
            }
        }

        /// <summary>
        /// Extantion methud Convers the Project XML file to ProjectModel
        /// </summary>
        /// <param name="xmlFile">Source file name</param>
        /// <returns>ProjectModel</returns>
        public static ProjectModel XmlProjectFileToModel(this string xmlFile)
        {
            ProjectModel p = new ProjectModel();
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
            XmlNodeList noteEmployee = xmlDoc.GetElementsByTagName("NoteEmployee");
            XmlNodeList projectNotes = xmlDoc.GetElementsByTagName("ProjectNotes");

            p.ProjectName = projectName[0].InnerText;
            p.NoteEmployee = bool.Parse(noteEmployee[0].InnerText);
            p.ProjectNotes = projectNotes[0].InnerText;
                
            return p;
        }

       /// <summary>
       /// Extantion methud Import the XML file to employee model
       /// </summary>
       /// <param name="xmlFile">Full path to XML file</param>
       /// <returns>EmployeeModel</returns>
        public static List<EmployeeModel> XmlEmloyeesFileToModel(this string xmlFile)
        {
            List<EmployeeModel> output = new List<EmployeeModel>();
            XmlDocument xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.Load(xmlFile);
            }
            catch ( Exception ex)
            {
                    
                MessageBox.Show(ex.Message, "SaveAsPDF:XmlEmloyeesFileToModel");
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
                output.Add(em);
            }

            //xmlDoc.Save(xmlFile);
            return output;
        }

        //public static List<string> XmlEmlpoyeesFileToList(this string xmlFile)
        //{
        //    List<EmployeeModel> e = new List<EmployeeModel>();
            
        //    List<string> output = new List<string>();
        //    XmlDocument xmlDoc = new XmlDocument();

        //    try
        //    {
        //        xmlDoc.Load(xmlFile);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "SaveAsPDF");
        //        return null;
        //    }

        //    XmlNodeList FirstName = xmlDoc.GetElementsByTagName("FirstName");
        //    XmlNodeList LastName = xmlDoc.GetElementsByTagName("LastName");
        //    XmlNodeList EmailAddress = xmlDoc.GetElementsByTagName("EmailAddress");

        //    for (int i = 0; i < FirstName.Count; i++)
        //    {
        //        EmployeeModel em = new EmployeeModel
        //        {
        //            FirstName = FirstName[i].InnerText,
        //            LastName = LastName[i].InnerText,
        //            EmailAddress = EmailAddress[i].InnerText
        //        };
        //        e.Add(em);
        //        output.Add(em.FullName);
        //    }
        //    return output;
        //}
        

    }
}
