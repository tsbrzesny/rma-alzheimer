using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Technewlogic.WpfDialogManagement;
using Technewlogic.WpfDialogManagement.Contracts;

namespace WPF_Roots
{

    public enum Buttons
    {
        None,
        Ok,
        Cancel,
        OkCancel,
        YesNo,
        YesNoCancel
    }


    public partial class MsgBox : UserControl
    {

        #region Static MsgBox factories

        public static ICustomContentDialog Info(string message, string title = null, object additionalContent = null,
                                                string message2 = null, Buttons dlgMode = Buttons.Ok)
        {
            return dMan.CreateCustomContentDialog(new MsgBox(message, title, additionalContent, message2, MsgBoxIcon.Info), null, (Technewlogic.WpfDialogManagement.DialogMode)dlgMode);
        }

        public static ICustomContentDialog Question(string message, string title = null, object additionalContent = null,
                                                    string message2 = null, Buttons dlgMode = Buttons.OkCancel)
        {
            return dMan.CreateCustomContentDialog(new MsgBox(message, title, additionalContent, message2, MsgBoxIcon.Question), null, (Technewlogic.WpfDialogManagement.DialogMode)dlgMode);
        }

        public static ICustomContentDialog Warning(string message, string title = null, object additionalContent = null,
                                                   string message2 = null, Buttons dlgMode = Buttons.OkCancel)
        {
            return dMan.CreateCustomContentDialog(new MsgBox(message, title, additionalContent, message2, MsgBoxIcon.Warning), null, (Technewlogic.WpfDialogManagement.DialogMode)dlgMode);
        }

        public static ICustomContentDialog Error(string message, string title = null, object additionalContent = null,
                                                 string message2 = null, Buttons dlgMode = Buttons.Cancel)
        {
            return dMan.CreateCustomContentDialog(new MsgBox(message, title, additionalContent, message2, MsgBoxIcon.Error), null, (Technewlogic.WpfDialogManagement.DialogMode)dlgMode);
        }

        # endregion

        #region Static members

        public static DialogManager dMan { get; private set; }

        public static void Init(ContentControl parent, System.Windows.Threading.Dispatcher dispatcher)
        {
            dMan = new DialogManager(parent, dispatcher);
        }

        #endregion



        public enum MsgBoxIcon
        {
            None,
            Warning,
            Error,
            Question,
            Info
        }

        Color cl_Neutral = Color.FromArgb(0xFF, 0xC0, 0xC0, 0xC0);
        Color cl_Info = Color.FromArgb(0xFF, 0x40, 0x40, 0xFF);
        Color cl_Warning = Color.FromArgb(0xFF, 0xFF, 0x68, 0x40);
        Color cl_Error = Color.FromArgb(0xFF, 0xFF, 0x40, 0x40);

        //
        public MsgBox(string message, string title = null, object additionalContent = null, string message2 = null,
                             MsgBoxIcon mbi = MsgBoxIcon.None)
        {
            InitializeComponent();

            // setup up dialog fields
            if (title == null)
                tb_title.Visibility = Visibility.Collapsed;
            else
            {
                tb_title.Visibility = Visibility.Visible;
                tb_title.Background = new SolidColorBrush(cl_Neutral);
                tb_title.Text = title;
            }

            tb_msg.Visibility = Visibility.Collapsed;
            if (message != null)
            {
                tb_msg.Visibility = Visibility.Visible;
                tb_msg.Text = message;
            }

            i_icon.Visibility = Visibility.Visible;
            switch (mbi)
            {
                case MsgBoxIcon.None:
                    i_icon.Visibility = Visibility.Collapsed;
                    break;
                case MsgBoxIcon.Question:
                    i_icon.Source = new BitmapImage(new Uri(@"Images/question.jpg", UriKind.Relative));
                    tb_title.Background = new SolidColorBrush(cl_Info);
                    break;
                case MsgBoxIcon.Info:
                    i_icon.Source = new BitmapImage(new Uri(@"Images/info.png", UriKind.Relative));
                    tb_title.Background = new SolidColorBrush(cl_Info);
                    break;
                case MsgBoxIcon.Warning:
                    i_icon.Source = new BitmapImage(new Uri(@"Images/warning.jpg", UriKind.Relative));
                    tb_title.Background = new SolidColorBrush(cl_Warning);
                    break;
                case MsgBoxIcon.Error:
                    i_icon.Source = new BitmapImage(new Uri(@"Images/error.png", UriKind.Relative));
                    tb_title.Background = new SolidColorBrush(cl_Error);
                    break;
            }

            cc_msg.Visibility = Visibility.Collapsed;
            if (cc_msg != null)
            {
                cc_msg.Visibility = Visibility.Visible;
                cc_msg.Content = additionalContent;
            }

            tb_msg2.Visibility = Visibility.Collapsed;
            if (message2 != null)
            {
                tb_msg2.Visibility = Visibility.Visible;
                tb_msg2.Text = message2;
            }
        }
    }
}
