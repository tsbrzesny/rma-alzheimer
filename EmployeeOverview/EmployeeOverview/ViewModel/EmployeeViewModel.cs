using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Windows.Media;
using System.Windows.Media.Converters;
using Alzheimer.Model;

namespace Alzheimer.ViewModel
{
    /// <summary>
    /// Employee View Model
    /// </summary>
    public class EmployeeViewModel : INotifyPropertyChanged
    {
        private IList<EmployeeModel> _employees;
        private string _searchTerm;
        private string _findings; 
        private ImageSource _employeeImage;
        private List<ImageSource> _employeeImageSources; 
         
        /// <summary>
        /// Employee View Model Constructor. 
        /// Will load all employees from excel
        /// </summary>
        public EmployeeViewModel()
        {
            Import import = new Import();
            _employees = import.Employees;
            _employeeImageSources = new List<ImageSource>();
            foreach (var employee in _employees)
            {
                _employeeImageSources.Add(employee.Image);
            }
        }


        public EmployeeModel AutoCompleteName(string partName)
        {
            var searchHitList = new List<EmployeeModel>(){};
            foreach (var em in _employees)
            {
                if (em.Zeichen.ToLower().Contains(partName.ToLower()))
                {
                    searchHitList.Add(em);
                }
            }
            if (searchHitList.Count == 1)
            {
                return searchHitList[0];
            }
            if (searchHitList.Count > 1)
            {
                return null;
            }
            if (searchHitList.Count == 0)
            {
                foreach (var em in _employees)
                {
                    if (em.FullName.ToLower().Contains(partName.ToLower()))
                    {
                        searchHitList.Add(em);
                    }
                }
            }
            if (searchHitList.Count == 1)
            {
                return searchHitList[0];
            }

            return null;
        }

        public List<ImageSource> EmployeeImageSources
        {
            get { return _employeeImageSources; }
            set
            {
                _employeeImageSources = value;
                NotifyPropertyChanged("EmployeeImageSources");
            }
        }

        public string SearchTerm
        {
            get { return _searchTerm; }
            set
            {
                _searchTerm = value;
                EmployeeModel eModel = AutoCompleteName(value);
                if (eModel != null)
                {
                    Findings = eModel.FullName;
                    EmployeeImage = eModel.Image;
                }
                else
                {
                    Findings = "";
                    EmployeeImage = null;
                }
                NotifyPropertyChanged("SearchTerm");
            }
        }

        public string Findings
        {
            get { return _findings; }
            set
            {
                _findings = value;
                NotifyPropertyChanged("Findings");
            }
        }

        public ImageSource EmployeeImage
        {
            get
            {
                return _employeeImage;
            }
            set
            {
                _employeeImage = value;
                NotifyPropertyChanged("EmployeeImage");
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }


    }
}
