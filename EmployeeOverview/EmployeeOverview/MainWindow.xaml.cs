using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Alzheimer.Model;
using Alzheimer.ViewModel;

namespace Alzheimer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            
            //DataContext = new EmployeeViewModel();
        }
    }
}