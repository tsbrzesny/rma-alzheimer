using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Alzheimer.Model;
using Alzheimer.ViewModel;


namespace Alzheimer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private Boolean _dragFlag;
        private Point _point;
        public MainWindow()
        {
            InitializeComponent();
            _dragFlag = false;
            RMA_Roots.AppConfig.LoadAppConfiguration("EmployeeOverview");
            DataContext = new EmployeeViewModel();
        }


        private void ListBox_OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            _dragFlag = true;
            _point = Mouse.GetPosition(this);
        }

        private void ListBox_Drop(object sender, DragEventArgs e)
        {
            if (_dragFlag)
            {
                var parent = (ListBox)sender;
                var data = e.Data.GetData(typeof(EmployeeModel));
                parent.Items.Add(data);
                _dragFlag = false;
                _point = new Point(0, 0); 
            }
        }

        private void UIElement_OnMouseMove(object sender, MouseEventArgs e)
        {
            var point = Mouse.GetPosition(this);
            if (_dragFlag)
            {
                if ((point.X + point.Y)+20 > (_point.X + _point.Y))
                {
                    DragDrop.DoDragDrop(ListBox1, ListBox1.SelectedItem, DragDropEffects.Move);
                }
            }
        }

        private void UIElement_OnMouseUp(object sender, MouseButtonEventArgs e)
        {
            _dragFlag = false;
            _point = new Point(0,0);
        }

        private void ListBox2_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListBox2.Items.Remove(ListBox2.SelectedItem);
        }
    }


}
