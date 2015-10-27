using System.Windows;
using Alzheimer.ViewModel;

namespace Alzheimer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        
        public MainWindow()
        {
            InitializeComponent();

            RMA_Roots.AppConfig.LoadAppConfiguration("EmployeeOverview");
            DataContext = new EmployeeViewModel();
        }
    }
}
