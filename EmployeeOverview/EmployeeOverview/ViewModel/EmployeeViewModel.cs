using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
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
        private EmployeeModel _findings; 
        private readonly ImageSource _sitzplanSource;
        private Visibility _showSeatingVisibility;
        
        
        /// <summary>
        /// Employee View Model Constructor. 
        /// Will load all employees from excel
        /// </summary>
        public EmployeeViewModel()
        {
            RMA_Roots.AppConfig.LoadAppConfiguration("EmployeeOverview");
            _sitzplanSource = new BitmapImage(new Uri(@"D:\sitzplan.png"));
            var import = new Import();
            _employees = import.Employees;
            _employees = _employees.OrderBy(f => f.Vorname).ToList();
            _showSeatingVisibility = Visibility.Hidden;
        }

        //TODO change this into something a little more appealing
        private EmployeeModel AutoCompleteName(string partName)
        {
            var searchHitList = new List<EmployeeModel>();
            
            //search through all employees.zeichen and try to find the match with the partName
            foreach (var em in _employees)
            {
                if (em.Zeichen.ToLower().Contains(partName.ToLower()))
                {
                    searchHitList.Add(em);
                }
            }
            
            //Zeichen of employee was found. Return found employee
            if (searchHitList.Count == 1)
                return searchHitList[0];
            
            //Meaning none Employee was unique identified to this point, so searching for name / prename
            foreach (var em in _employees)
            {
                if (em.FullName.ToLower().Contains(partName.ToLower()))
                {
                    searchHitList.Add(em);
                }
            }
            
            //If just one was found -> display the found one. Otherwise return null
            if (searchHitList.Count == 1)
            {
                return searchHitList[0];
            }

            return null;
        }


        /// <summary>
        /// Visibility of the seating.
        /// should only be visible when an employee is displayed
        /// </summary>
        public Visibility ShowSeatingVisibility
        {
            get { return _showSeatingVisibility; }
            set
            {
                _showSeatingVisibility = value;
                NotifyPropertyChanged("ShowSeatingVisibility");
            }
        }

        /// <summary>
        /// Picture of the Sitzplan
        /// </summary>
        public ImageSource SitzplanSource
        {
            get { return _sitzplanSource;}
        }

        /// <summary>
        /// Searchterm
        /// </summary>
        public string SearchTerm
        {
            get { return _searchTerm; }
            set
            {
                _searchTerm = value;
                EmployeeModel eModel = AutoCompleteName(value);
                if (eModel != null)
                {
                    Findings = eModel;
                    //EmployeeImage = eModel.Image;
                }
                else
                {
                    Findings = null;
                    //EmployeeImage = null;
                    ShowSeatingVisibility = Visibility.Hidden;
                }
                NotifyPropertyChanged("SearchTerm");
            }
        }

        /// <summary>
        /// Results of the search
        /// </summary>
        public EmployeeModel Findings
        {
            get { return _findings; }
            set
            {
                _findings = value;
                if (value != null)
                {
                    ShowSeatingVisibility = Visibility.Visible;
                }
                NotifyPropertyChanged("Findings");
            }
        }

        /// <summary>
        /// List of all employee's
        /// </summary>
        public IList<EmployeeModel> Employees
        {
            get { return _employees; }
            set { _employees = value; }
        }

        public IList<EmployeeModel> FavoriteEmployees { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Notifiy when a property changed so the view is updating its bindings
        /// </summary>
        /// <param name="propertyName">Name of the Property that needs an update</param>
        private void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }


    }
}
