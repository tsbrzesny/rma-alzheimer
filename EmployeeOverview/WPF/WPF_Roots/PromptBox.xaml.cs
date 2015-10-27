using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace WPF_Roots
{
    /// <summary>
    /// Interaction logic for PromptBox.xaml
    /// </summary>
    public partial class PromptBox : Window
    {
        public PromptBox(string title, string message, object promptContent)
        {
            InitializeComponent();

            this.tb_title.Text = title;
            this.tb_msg.Text = message;
            this.cc_msg.Content = promptContent;
        }

        public enum PromptBoxIcon
        {
            None,
            Warning,
            Error,
            Question,
            Info,
            Prompt,
            Logout,
        }

        Color cl_Neutral = Color.FromArgb(0xFF, 0xC0, 0xC0, 0xC0);
        Color cl_Info = Color.FromArgb(0xFF, 0x40, 0x40, 0xFF);
        Color cl_Warning = Color.FromArgb(0xFF, 0xFF, 0x68, 0x40);
        Color cl_Error = Color.FromArgb(0xFF, 0xFF, 0x40, 0x40);

        //
        public PromptBox(string title, object content, PromptBoxIcon iconType = PromptBoxIcon.None)
        {
            InitializeComponent();

            // setup up dialog fields
            tb_title.Visibility = Visibility.Collapsed;
            if (title != null)
            {
                tb_title.Visibility = Visibility.Visible;
                tb_title.Text = title;
                tb_title.Background = new SolidColorBrush(cl_Neutral);
            }

            tb_msg.Visibility = Visibility.Collapsed;
            frame_buttons.Visibility = Visibility.Collapsed;
            cc_msg.Content = content;

            switch (iconType)
            {
                case PromptBoxIcon.None:
                    i_icon.Visibility = Visibility.Collapsed;
                    break;
                case PromptBoxIcon.Question:
                    i_icon.Source = new BitmapImage(new Uri(@"Images/question.jpg", UriKind.Relative));
                    tb_title.Background = new SolidColorBrush(cl_Info);
                    break;
                case PromptBoxIcon.Info:
                    i_icon.Source = new BitmapImage(new Uri(@"Images/info.png", UriKind.Relative));
                    tb_title.Background = new SolidColorBrush(cl_Info);
                    break;
                case PromptBoxIcon.Warning:
                    i_icon.Source = new BitmapImage(new Uri(@"Images/warning.jpg", UriKind.Relative));
                    tb_title.Background = new SolidColorBrush(cl_Warning);
                    break;
                case PromptBoxIcon.Error:
                    i_icon.Source = new BitmapImage(new Uri(@"Images/error.png", UriKind.Relative));
                    tb_title.Background = new SolidColorBrush(cl_Error);
                    break;
                case PromptBoxIcon.Prompt:
                    i_icon.Source = new BitmapImage(new Uri(@"Images/prompt.png", UriKind.Relative));
                    tb_title.Background = new SolidColorBrush(cl_Info);
                    break;
                case PromptBoxIcon.Logout:
                    i_icon.Source = new BitmapImage(new Uri(@"Images/logout.png", UriKind.Relative));
                    break;
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
        }

        public bool DialogCancelled = false;
        //
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogCancelled = true;
                Close();
                e.Handled = true;
            }
        }

        private void bt_Ok_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

    }
}
