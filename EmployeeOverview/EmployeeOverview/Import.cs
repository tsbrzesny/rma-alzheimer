using System;
using System.Collections.Generic;
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
        }

        private void ImportDb()
        {
            _employees = new List<EmployeeModel>();
            foreach (RMA2D.RMAMitarbeiter ma in RMA2D.mitarbeiterT)
            {
                EmployeeModel eModel = new EmployeeModel(ma);
                _employees.Add(eModel);
            }
        }


        #region helper
        /// <summary>
        /// if string contains "''" replace it with ""
        /// </summary>
        /// <param name="text">text to check</param>
        /// <returns>string without any "''"</returns>
        private String ToStringEmpty(Object text)
        {
            return text.ToString().Equals("''") ? "" : text.ToString();
        }
        #endregion



        public List<EmployeeModel> Employees
        {
            get { return _employees; }
            set { _employees = value; }
        }
    }
}
