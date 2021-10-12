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

namespace SaveAsPDF.Helpers
{
   public static class XmlFileHelper
    {
       
        public static void ProjectModelToXmlFile(ProjectModel project, string path)
        {

            
            XDocument xmlDoc = new XDocument(
                                        new XElement("Project", (
                                            new XElement("ProjectNumber", project.ProjectNumber),
                                            new XElement("ProjectName", project.ProjectName),
                                            new XElement("NoteEmployee", project.NoteEmployee))));
                                
            xmlDoc.Save(path);
            
        }

        public static void EmployeesModelToXmlFile(List<EmployeeModel> employees, string path)
        {
            XDocument Output = new XDocument();

            foreach (EmployeeModel e in employees)
            {
                XDocument xmlDoc =  new XDocument(
                                                new XElement("Employees", (
                                                    new XElement("Employee", (
                                                    new XElement("ID", e.Id.ToString()),
                                                    new XElement("FirstName", e.FirstName)),
                                                    new XElement("LastName", e.LastName)),
                                                    new XElement("EmailAddress", e.EmailAddress))));
                Output.Add(xmlDoc);
            }
            Output.Save(path);

        }



        //public static void CreateProjectXmlFile(string path, ProjectModel project)
        //{

        //    XmlWriter xmlWriter = XmlWriter.Create($"{path}{GlobalConfig.xmlProjectFile}");

        //    xmlWriter.WriteStartDocument();
        //    xmlWriter.WriteStartElement("Project");

        //    xmlWriter.WriteStartElement("ProjectNumber");
        //    xmlWriter.WriteString(project.ProjectNumber);
        //    xmlWriter.WriteEndElement();

        //    xmlWriter.WriteStartElement("ProjectName");
        //    xmlWriter.WriteString(project.ProjectName);
        //    xmlWriter.WriteEndElement();

        //    xmlWriter.WriteStartElement("NoteEmployee");
        //    xmlWriter.WriteString(project.NoteEmployee.ToString());
        //    xmlWriter.WriteEndElement();

        //    xmlWriter.WriteEndElement();
        //    xmlWriter.WriteEndDocument();
        //    xmlWriter.Close();

        //}

        //public static void CreateEmployeesXmlFile(string path, EmployeeModel employee)
        //{

        //    XmlWriter xmlWriter = XmlWriter.Create($"{path}{GlobalConfig.xmlEmploeeysFile}");

        //    xmlWriter.WriteStartDocument();
        //    xmlWriter.WriteStartElement("Employees");

        //    xmlWriter.WriteStartElement("Employee");
        //    xmlWriter.WriteString(employee.FirstName);
        //    xmlWriter.WriteEndElement();

        //    xmlWriter.WriteStartElement("ProjectName");
        //    xmlWriter.WriteString(employee.LastName);
        //    xmlWriter.WriteEndElement();

        //    xmlWriter.WriteStartElement("NoteEmployee");
        //    xmlWriter.WriteString(employee.EmailAddress);
        //    xmlWriter.WriteEndElement();

        //    xmlWriter.WriteEndElement();
        //    xmlWriter.WriteEndDocument();
        //    xmlWriter.Close();

        //}



        public static ProjectModel XmlProjectFileToModel(this string xmlFile)
        {
            ProjectModel p = new ProjectModel();
            XmlDocument xmlDoc = new XmlDocument();

            if (File.Exists(xmlFile))
            {
                xmlDoc.Load(xmlFile);
                XmlNodeList projectName = xmlDoc.GetElementsByTagName("ProjectName");
                XmlNodeList NoteEmployee = xmlDoc.GetElementsByTagName("NoteEmployee");

                p.ProjectName = projectName[0].InnerText;
                p.NoteEmployee = bool.Parse(NoteEmployee[0].InnerText);
            }
            return p;
        }

       /// <summary>
       /// Import the XML file to employee model
       /// </summary>
       /// <param name="xmlFile"></param>
       /// <returns>EmployeeModel</returns>
        public static List<EmployeeModel> XmlEmloyeesFileToModel(this string xmlFile)
        {
            List<EmployeeModel> output = new List<EmployeeModel>();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFile);

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
            return output;
        }



        public static List<string> XmlEmlpoyeesFileToList(this string xmlFile)
        {
            List<EmployeeModel> e = new List<EmployeeModel>();
            
            List<string> output = new List<string>();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFile);

            XmlNodeList FirstName = xmlDoc.GetElementsByTagName("FirstName");
            XmlNodeList LastName = xmlDoc.GetElementsByTagName("LastName");
            XmlNodeList EmailAddress = xmlDoc.GetElementsByTagName("EmailAddress");

            for (int i = 0; i < FirstName.Count; i++)
            {
                EmployeeModel em = new EmployeeModel
                {
                    FirstName = FirstName[i].InnerText,
                    LastName = LastName[i].InnerText,
                    EmailAddress = EmailAddress[i].InnerText
                };
                e.Add(em);
                output.Add(em.FullName);
            }
            return output;
        }
        

    }
}
