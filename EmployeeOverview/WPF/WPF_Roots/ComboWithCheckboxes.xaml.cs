﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WPF_Roots
{
    /// <summary>
    /// Interaction logic for ComboWithCheckboxes.xaml
    /// </summary>
    public partial class ComboWithCheckboxes
    {

        #region Dependency Properties

        /// <summary>
        ///Gets or sets a collection used to generate the content of the ComboBox
        /// </summary>
        public object ItemsSource
        {
            get { return (object)GetValue(ItemsSourceProperty); }
            set
            {
                SetValue(ItemsSourceProperty, value);
                SetText();
            }
        }

        public static readonly DependencyProperty ItemsSourceProperty =
            DependencyProperty.Register("ItemsSource", typeof(object), typeof(ComboWithCheckboxes), new UIPropertyMetadata(null));

        /// <summary>
        ///Gets or sets the text displayed in the ComboBox
        /// </summary>
        public string Text
        {
            get { return(string)GetValue(TextProperty); }
            set{ SetValue(TextProperty, value); }
        }

        public static readonly DependencyProperty TextProperty =
            DependencyProperty.Register("Text", typeof(string), typeof(ComboWithCheckboxes), new UIPropertyMetadata(string.Empty));

        /// <summary>
        ///Gets or sets the text displayed in the ComboBox if there are no selected items
        /// </summary>
        public string DefaultText
        {
            get { return (string)GetValue(DefaultTextProperty); }
            set { 
                SetValue(DefaultTextProperty, value);
                SetText();
            }
        }

        // Using a DependencyProperty as the backing store for FixText.  This enables animation, styling, binding, etc…
        public static readonly DependencyProperty DefaultTextProperty =
            DependencyProperty.Register("DefaultText", typeof(string), typeof(ComboWithCheckboxes), new UIPropertyMetadata(string.Empty));


        /// <summary>
        ///Gets or sets the text displayed in the ComboBox if there are no selected items
        /// </summary>
        public string FixText
        {
            get { return (string)GetValue(FixTextProperty); }
            set { 
                SetValue(FixTextProperty, value);
                SetText();
            }
        }

        // Using a DependencyProperty as the backing store for FixText.  This enables animation, styling, binding, etc…
        public static readonly DependencyProperty FixTextProperty =
            DependencyProperty.Register("FixText", typeof(string), typeof(ComboWithCheckboxes), new UIPropertyMetadata(string.Empty));

        #endregion



        public ComboWithCheckboxes()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Whenever a CheckBox is checked, change the text displayed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            SetText();
        }

        /// <summary>
        ///Set the text property of this control (bound to the ContentPresenter of the ComboBox)
        /// </summary>

        private void SetText()
        {
            if (!string.IsNullOrEmpty(this.FixText))
            {
                this.Text = this.FixText;
            }

            else
            {
                this.Text = (this.ItemsSource != null) ?
                this.ItemsSource.ToString() : this.DefaultText;

                // set DefaultText if nothing else selected
                if (string.IsNullOrEmpty(this.Text))
                    this.Text = this.DefaultText;
            }
        }

    }
}
