using System;
using System.Collections.Generic;
using System.Xml;
using Alzheimer.Model;
using RMA2_Roots;



namespace Alzheimer
{
    /// <summary>
    /// Import class with different mechanisms for importing employees of RMA
    /// </summary>
    public class Import
    {
        private List<EmployeeModel> _employees;

        public Import()
        {
            ImportDb();
            ImportXML();
        }

        private void ImportXML()
        {
            XmlDocument doc = new XmlDocument();
            doc.Load("d:\\Mitarbeiter.xml");
            if (doc.DocumentElement != null)
            {
                XmlNodeList nodes = doc.DocumentElement.SelectNodes("Mitarbeiter");
                foreach (XmlNode node in nodes)
                {
                    var myName = node.ChildNodes[0].InnerText;
                    var myX = node.ChildNodes[1].InnerText;
                    var myY = node.ChildNodes[2].InnerText;

                    for (int i = 0; i < _employees.Count; i++)
                    {
                        if (_employees[i].FullName.Equals(myName))
                        {
                            _employees[i].PosX = Int32.Parse(myX);
                            _employees[i].PosY = Int32.Parse(myY);
                        }
                    }
                }
            }
        }

        private void ImportDb()
        {
            _employees = new List<EmployeeModel>();
            foreach (var ma in RMA2D.mitarbeiterT)
            {
                var eModel = new EmployeeModel(ma);
                _employees.Add(eModel);
            }
        }

        public List<EmployeeModel> Employees
        {
            get { return _employees; }
        }
    }
}
