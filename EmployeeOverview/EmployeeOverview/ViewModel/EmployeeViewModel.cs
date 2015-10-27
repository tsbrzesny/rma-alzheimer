using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Media;
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
        private EmployeeModel _selectedEmployee;
        private IList<EmployeeModel> _favoriteEmployees; 
         
        /// <summary>
        /// Employee View Model Constructor. 
        /// Will load all employees from excel
        /// </summary>
        public EmployeeViewModel()
        {
            var import = new Import();
            _employees = import.Employees;
            _employees = _employees.OrderBy(f => f.Vorname).ToList();
        }


        public EmployeeModel AutoCompleteName(string partName)
        {
            var searchHitList = new List<EmployeeModel>();
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

        

        public string SearchTerm
        {
            get { return _searchTerm; }
            set
            {
                _searchTerm = value;
                EmployeeModel eModel = AutoCompleteName(value);
                if (eModel != null)
                {
                    Findings = eModel.FullName + " " + eModel.ContactId;
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

        public EmployeeModel SelectedEmployee
        {
            get { return _selectedEmployee; }
            set
            {
                _selectedEmployee = value;
                SearchTerm = _selectedEmployee.FullName;
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

        public IList<EmployeeModel> Employees
        {
            get { return _employees; }
            set { _employees = value; }
        }

        public IList<EmployeeModel> FavoriteEmployees
        {
            get { return _favoriteEmployees; }
            set { _favoriteEmployees = value; }
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
