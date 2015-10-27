using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using System.Xml.Linq;
using System.Xml.XPath;

namespace WPF_Roots
{

    /// <summary>
    /// Helper methods for UI-related tasks.
    /// from http://www.hardcodet.net/uploads/2009/06/UIHelper.cs
    /// </summary>
    public static class UIHelpers
    {
        #region find parent

        /// <summary>
        /// Finds a parent of a given item on the visual tree.
        /// </summary>
        /// <typeparam name="T">The type of the queried item.</typeparam>
        /// <param name="child">A direct or indirect child of the
        /// queried item.</param>
        /// <returns>The first parent item that matches the submitted
        /// type parameter. If not matching item can be found, a null
        /// reference is being returned.</returns>
        public static T TryFindParent<T>(this DependencyObject child)
            where T : DependencyObject
        {
            //get parent item
            DependencyObject parentObject = GetParentObject(child);

            //we've reached the end of the tree
            if (parentObject == null) return null;

            //check if the parent matches the type we're looking for
            T parent = parentObject as T;
            if (parent != null)
            {
                return parent;
            }
            else
            {
                //use recursion to proceed with next level
                return TryFindParent<T>(parentObject);
            }
        }

        /// <summary>
        /// This method is an alternative to WPF's
        /// <see cref="VisualTreeHelper.GetParent"/> method, which also
        /// supports content elements. Keep in mind that for content element,
        /// this method falls back to the logical tree of the element!
        /// </summary>
        /// <param name="child">The item to be processed.</param>
        /// <returns>The submitted item's parent, if available. Otherwise
        /// null.</returns>
        public static DependencyObject GetParentObject(this DependencyObject child)
        {
            if (child == null) return null;

            //handle content elements separately
            ContentElement contentElement = child as ContentElement;
            if (contentElement != null)
            {
                DependencyObject parent = ContentOperations.GetParent(contentElement);
                if (parent != null) return parent;

                FrameworkContentElement fce = contentElement as FrameworkContentElement;
                return fce != null ? fce.Parent : null;
            }

            //also try searching for parent in framework elements (such as DockPanel, etc)
            FrameworkElement frameworkElement = child as FrameworkElement;
            if (frameworkElement != null)
            {
                DependencyObject parent = frameworkElement.Parent;
                if (parent != null) return parent;
            }

            //if it's not a ContentElement/FrameworkElement, rely on VisualTreeHelper
            return VisualTreeHelper.GetParent(child);
        }

        #endregion

        #region find children

        /// <summary>
        /// Analyzes both visual and logical tree in order to find all elements of a given
        /// type that are descendants of the <paramref name="source"/> item.
        /// </summary>
        /// <typeparam name="T">The type of the queried items.</typeparam>
        /// <param name="source">The root element that marks the source of the search. If the
        /// source is already of the requested type, it will not be included in the result.</param>
        /// <returns>All descendants of <paramref name="source"/> that match the requested type.</returns>
        public static IEnumerable<T> FindChildren<T>(this DependencyObject source) where T : DependencyObject
        {
            if (source != null)
            {
                var childs = GetChildObjects(source);
                foreach (DependencyObject child in childs)
                {
                    //analyze if children match the requested type
                    if (child != null && child is T)
                    {
                        yield return (T)child;
                    }

                    //recurse tree
                    foreach (T descendant in FindChildren<T>(child))
                    {
                        yield return descendant;
                    }
                }
            }
        }


        /// <summary>
        /// This method is an alternative to WPF's
        /// <see cref="VisualTreeHelper.GetChild"/> method, which also
        /// supports content elements. Keep in mind that for content elements,
        /// this method falls back to the logical tree of the element.
        /// </summary>
        /// <param name="parent">The item to be processed.</param>
        /// <returns>The submitted item's child elements, if available.</returns>
        public static IEnumerable<DependencyObject> GetChildObjects(this DependencyObject parent)
        {
            if (parent == null) yield break;

            if (parent is Visual || parent is Visual3D)
            {
                //use the visual tree per default
                int count = VisualTreeHelper.GetChildrenCount(parent);
                for (int i = 0; i < count; i++)
                {
                    yield return VisualTreeHelper.GetChild(parent, i);
                }
            }
            else
            {
                //use the logical tree for content / framework elements
                foreach (object obj in LogicalTreeHelper.GetChildren(parent))
                {
                    var depObj = obj as DependencyObject;
                    if (depObj != null) yield return (DependencyObject)obj;
                }
            }
        }

        #endregion


        #region find from point

        /// <summary>
        /// Tries to locate a given item within the visual tree,
        /// starting with the dependency object at a given position. 
        /// </summary>
        /// <typeparam name="T">The type of the element to be found
        /// on the visual tree of the element at the given location.</typeparam>
        /// <param name="reference">The main element which is used to perform
        /// hit testing.</param>
        /// <param name="point">The position to be evaluated on the origin.</param>
        public static T TryFindFromPoint<T>(UIElement reference, Point point)
            where T : DependencyObject
        {
            DependencyObject element = reference.InputHitTest(point) as DependencyObject;

            if (element == null) return null;
            else if (element is T) return (T)element;
            else return TryFindParent<T>(element);
        }

        #endregion


        #region Points, Rects

        public static Point Center(this Rect rect)
        {
            return new Point(rect.Left + rect.Width / 2,
                             rect.Top + rect.Height / 2);
        }


        public static double Distance(this Point p1, Point p2, bool squareIsOk = false)
        {
            double dist = Math.Pow(p1.X - p2.X, 2) + Math.Pow(p1.Y - p2.Y, 2);
            if (squareIsOk)
                return dist;
            else
                return Math.Sqrt(dist);
        }

        # endregion

    }


    public class AttachedProperties
    {
        /*
        public static readonly DependencyProperty SortInstructionsProperty = DependencyProperty.RegisterAttached(
                "SortInstructions", typeof(string), typeof(AttachedProperties), new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.None));
        public static string GetSortInstructions(DependencyObject obj)
        {
            return (string)obj.GetValue(SortInstructionsProperty);
        }
        public static void SetSortInstructions(DependencyObject obj, string value)
        {
            obj.SetValue(SortInstructionsProperty, value);
        }
        */
    }

    public class ListViewGridViewHelper
    {
        public static readonly DependencyProperty SortInstructionsProperty = DependencyProperty.RegisterAttached(
                "SortInstructions", typeof(string), typeof(ListViewGridViewHelper), new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.None));
        public static string GetSortInstructions(DependencyObject obj)
        {
            return (string)obj.GetValue(SortInstructionsProperty);
        }
        public static void SetSortInstructions(DependencyObject obj, string value)
        {
            obj.SetValue(SortInstructionsProperty, value);
        }

        public static readonly DependencyProperty LastColumnSortedProperty = DependencyProperty.RegisterAttached(
                "LastColumnSorted", typeof(string), typeof(ListViewGridViewHelper), new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.None));
        public static string GetLastColumnSorted(DependencyObject obj)
        {
            return (string)obj.GetValue(LastColumnSortedProperty);
        }
        public static void SetLastColumnSorted(DependencyObject obj, string value)
        {
            obj.SetValue(LastColumnSortedProperty, value);
        }

        public static readonly DependencyProperty LastSortDirectionProperty = DependencyProperty.RegisterAttached(
                "LastSortDirection", typeof(ListSortDirection?), typeof(ListViewGridViewHelper), new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.None));
        public static ListSortDirection? GetLastSortDirection(DependencyObject obj)
        {
            return (ListSortDirection?)obj.GetValue(LastSortDirectionProperty);
        }
        public static void SetLastSortDirection(DependencyObject obj, ListSortDirection? value)
        {
            obj.SetValue(LastSortDirectionProperty, value);
        }

    }


    public static class ListViewExtensions
    {
        // expects a sortInstructions string like "↓GroupKey↑Group2↕Item"
        // ↓ = sort ascending (fix), ↑ = sort descending (fix), ↕ = column sorting (toggleable)
        //
        public static ListSortDirection EstablishCodedSortOrder(this ListView listView, string sortInstructions, ListSortDirection? overrideToggleDirection = null )
        {
            var dataView = CollectionViewSource.GetDefaultView(listView.ItemsSource);
            var establishedColumnDirection = ListSortDirection.Ascending;

            if (dataView != null && dataView.SortDescriptions != null && sortInstructions != null)
            {
                dataView.SortDescriptions.Clear();
                foreach (Match m in Regex.Matches(sortInstructions, @"([↓↑↕])([^↓↑↕]+)"))
                {
                    var dirCode = m.Groups[1].Value;
                    var fieldCode = m.Groups[2].Value;
                    if (dirCode == "↓")
                        dataView.SortDescriptions.Add(new SortDescription(fieldCode, ListSortDirection.Ascending));
                    else if (dirCode == "↑")
                        dataView.SortDescriptions.Add(new SortDescription(fieldCode, ListSortDirection.Descending));
                    else
                    {
                        string lastColumnSorted = ListViewGridViewHelper.GetLastColumnSorted(listView);
                        ListSortDirection? lastSortDirection = ListViewGridViewHelper.GetLastSortDirection(listView);

                        if (overrideToggleDirection.HasValue)
                            establishedColumnDirection = overrideToggleDirection.Value;

                        else if (lastColumnSorted != null && fieldCode == lastColumnSorted && lastSortDirection.HasValue && lastSortDirection.Value == ListSortDirection.Ascending)
                            establishedColumnDirection = ListSortDirection.Descending;

                        dataView.SortDescriptions.Add(new SortDescription(fieldCode, establishedColumnDirection));
                        ListViewGridViewHelper.SetLastColumnSorted(listView, fieldCode);
                        ListViewGridViewHelper.SetLastSortDirection(listView, establishedColumnDirection);
                    }
                }
                dataView.Refresh();
            }
            return establishedColumnDirection;
        }

    }



    public class ControlKeyStates
    {
        public bool shiftState { get; private set; }
        public bool ctrlState { get; private set; }
        public bool altState { get; private set; }

        public ControlKeyStates()
        {
            UpdateControlKeyStates();
        }

        public void UpdateControlKeyStates()
        {
            shiftState = Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);
            ctrlState = Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);
            altState = Keyboard.IsKeyDown(Key.LeftAlt) || Keyboard.IsKeyDown(Key.RightAlt);
        }

        public bool NoCtrlKey { get { return !shiftState && !ctrlState && !altState; } }

        public bool ShiftOnly { get { return shiftState && !ctrlState && !altState; } }
        public bool CtrlOnly { get { return !shiftState && ctrlState && !altState; } }
        public bool AltOnly { get { return !shiftState && !ctrlState && altState; } }

        public bool CtrlShift { get { return shiftState && ctrlState && !altState; } }

    }



    public static class GenericExtensions
    {

        #region List<>: Push, Pop, Shift, Unshift, IsEmpty

        public static void Push<T>(this List<T> aList, T newElement)
        {
            aList.Add(newElement);
        }

        public static T Pop<T>(this List<T> aList)
        {
            if (aList.Count > 0)
            {
                var lastIndex = aList.Count - 1;
                var lastElement = aList[lastIndex];
                aList.RemoveAt(lastIndex);
                return lastElement;
            }
            throw new Exception("Can't Pop, list is empty!");
        }

        public static void Unshift<T>(this List<T> aList, T newElement)
        {
            aList.Insert(0, newElement);
        }

        public static T Shift<T>(this List<T> aList)
        {
            if (aList.Count > 0)
            {
                var firstElement = aList.First();
                aList.RemoveAt(0);
                return firstElement;
            }
            throw new Exception("Can't Shift, list is empty!");
        }

        public static bool IsEmpty<T>(this List<T> aList)
        {
            return aList.Count == 0;
        }

        #endregion  // List<>


        #region XNode_XPath_Readers

        public static List<string> XPath_GetItems(this XNode xNode, string xpath)
        {
            // must import System.Xml.XPath
            IEnumerable value = (IEnumerable)xNode.XPathEvaluate(xpath);

            var xAttrs = value.Cast<XAttribute>();
            if (xAttrs != null)
                return xAttrs.Select(x => x.Value).ToList();

            var xElems = value.Cast<XElement>();
            if (xElems != null)
                return xElems.Select(x => x.Value).ToList();

            var result = new List<string>();
            if (value != null)
                result.Add(value.ToString());
            return result;
        }

        public static string XPath_GetFirstItem(this XNode xNode, string xpath)
        {
            var values = xNode.XPath_GetItems(xpath);
            if (values.Count() > 0)
                return values[0];
            return null;
        }

        public static string XPath_GetString(this XNode xNode, string xpath, string defaultValue = null)
        {
            return xNode.XPath_GetFirstItem(xpath) ?? defaultValue;
        }

        public static string XPath_GetFloat(this XNode xNode, string xpath, float defaultValue = float.NaN)
        {
            float result = defaultValue;

            string value = xNode.XPath_GetFirstItem(xpath);
            float.TryParse(value, out result);
            return value;
        }

        public static int XPath_GetInteger(this XNode xNode, string xpath, int defaultValue = int.MinValue)
        {
            int result;

            string value = xNode.XPath_GetFirstItem(xpath);
            if (value == null || !int.TryParse(value, out result))
                return defaultValue;
            return result;
        }

        public static bool XPath_GetBoolean(this XNode xNode, string xpath, 
                                            bool defaultValue = false, bool existenceMeansTrue = true)
        {
            string value = xNode.XPath_GetFirstItem(xpath);
            if (value == null)
                return defaultValue;         // element does not exist

            if (value.Length == 0)
                return existenceMeansTrue;   // element exists, but is empty

            value = value.ToUpper();
            if (value == "TRUE" || value == "T" || value == "1")
                return true;
            if (value == "False" || value == "F" || value == "0")
                return false;
            return defaultValue;
        }

        public static bool XPath_Exists(this XNode xNode, string xpath)
        {
            return (xNode.XPath_GetFirstItem(xpath) != null);
        }


        #endregion  // XNode_XPath_Readers


        #region JObject readers


        public static string GetString(this JObject jo, string key, string defaultValue = null)
        {
            try
            {
                return (string)jo[key];
            }
            catch (Exception)
            {
            }
            return defaultValue;
        }

        public static float GetFloat(this JObject jo, string key,  float defaultValue = float.NaN)
        {
            try
            {
                return float.Parse((string)jo[key]);
            }
            catch (Exception)
            {
            }
            return defaultValue;
        }

        public static int GetInt(this JObject jo, string key, int defaultValue = int.MinValue)
        {
            try
            {
                return int.Parse((string)jo[key]);
            }
            catch (Exception)
            {
            }
            return defaultValue;
        }

        public static bool GetBool(this JObject jo, string key, bool defaultValue = false, bool existenceMeansTrue = true)
        {
            try
            {
                var bStr = (string)jo[key];
                bStr = bStr.ToUpper();
                if (bStr == "TRUE" || bStr == "T" || bStr == "1")
                    return true;
                if (bStr == "False" || bStr == "F" || bStr == "0")
                    return false;
                return existenceMeansTrue;
            }
            catch (Exception)
            {
            }
            return defaultValue;
        }

        public static DateTime? GetDate(this JObject jo, string key, DateTime? defaultValue = null )
        {
            try
            {
                return (DateTime)jo[key];
            }
            catch (Exception)
            {
            }
            return defaultValue;
        }

        #endregion  // JObject readers


        #region String extensions

        public static Stream ToStream(this string s)
        {
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }

        public static string ToNull(this string s)
        {
            if (s == null || s.Length == 0)
                return null;
            return s;
        }

        public static string ToEmpty(this string s)
        {
            if (s == null)
                return "";
            return s;
        }

        public static bool IsNullOrEmpty(this string s)
        {
            return s == null || s.Length == 0;
        }

        public static bool IsLike(this string s, string pattern)
        {
            return Regex.IsMatch(s, pattern);
        }

        public static List<string> ListMatches(this string s, string pattern)
        {
            var result = new List<string>();
            try
            {
                var ms = Regex.Matches(s, pattern);
                foreach (Match m in ms)
                {
                    for (var gNr = 1; gNr <= m.Groups.Count; gNr++)
                        result.AddRange(from Capture c in m.Groups[gNr].Captures select c.Value);
                }
            }
            catch { }

            return result;
        }

        public static string NoWhitespace(this string s)
        {
            if (s == null)
                return null;
            return Regex.Replace(s, @"\s+", "");
        }

        public static string CleanWhitespace(this string s)
        {
            if (s == null)
                return null;
            return Regex.Replace(s, @"\s+", " ").Trim();
        }

        public static string GetEncodedHash(this string s, string salt = "")
        {
            if (s == null)
                return null;

            MD5 md5 = new MD5CryptoServiceProvider();
            byte [] digest = md5.ComputeHash(Encoding.UTF8.GetBytes(s + (salt ?? "")));
            return Convert.ToBase64String(digest, 0, digest.Length);
        }

        #endregion // String extensions


        // Evaluates the given ListView's AttachedProperties.SortInstructions and calls EstablishCodedSortOrder
        public static void EstablishInitialSortOrder(this ListView listView, string sortInstructions = null)
        {
            if (sortInstructions == null)
                sortInstructions = ListViewGridViewHelper.GetSortInstructions(listView);
            listView.EstablishCodedSortOrder(sortInstructions, ListSortDirection.Ascending);
        }

    }


    #region OcPropertyChangedListener
    public class OcPropertyChangedListener<T> : INotifyPropertyChanged where T : INotifyPropertyChanged
    {
        private readonly ObservableCollection<T> _collection;
        private readonly string _propertyName;
        private readonly Dictionary<T, int> _items = new Dictionary<T, int>(new ObjectIdentityComparer());
        public OcPropertyChangedListener(ObservableCollection<T> collection, string propertyName = "")
        {
            _collection = collection;
            _propertyName = propertyName ?? "";
            AddRange(collection);

            // CollectionChangedEventManager.AddHandler(collection, CollectionChanged); seems not 2 work.. :(
            collection.CollectionChanged += CollectionChanged;      // using strong reference instead
        }

        private void CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            switch (e.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    AddRange(e.NewItems.Cast<T>());
                    break;
                case NotifyCollectionChangedAction.Remove:
                    RemoveRange(e.OldItems.Cast<T>());
                    break;
                case NotifyCollectionChangedAction.Replace:
                    AddRange(e.NewItems.Cast<T>());
                    RemoveRange(e.OldItems.Cast<T>());
                    break;
                case NotifyCollectionChangedAction.Move:
                    break;
                case NotifyCollectionChangedAction.Reset:
                    Reset();
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

        }

        private void AddRange(IEnumerable<T> newItems)
        {
            foreach (T item in newItems)
            {
                if (_items.ContainsKey(item))
                {
                    _items[item]++;
                }
                else
                {
                    _items.Add(item, 1);
                    PropertyChangedEventManager.AddHandler(item, ChildPropertyChanged, _propertyName);
                }
            }
        }

        private void RemoveRange(IEnumerable<T> oldItems)
        {
            foreach (T item in oldItems)
            {
                _items[item]--;
                if (_items[item] == 0)
                {
                    _items.Remove(item);
                    PropertyChangedEventManager.RemoveHandler(item, ChildPropertyChanged, _propertyName);
                }
            }
        }

        private void Reset()
        {
            foreach (T item in _items.Keys.ToList())
            {
                PropertyChangedEventManager.RemoveHandler(item, ChildPropertyChanged, _propertyName);
                _items.Remove(item);
            }
            AddRange(_collection);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void ChildPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
                handler(sender, e);
        }

        private class ObjectIdentityComparer : IEqualityComparer<T>
        {
            public bool Equals(T x, T y)
            {
                return object.ReferenceEquals(x, y);
            }
            public int GetHashCode(T obj)
            {
                return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
            }
        }
    }

    public static class OcPropertyChangedListener
    {
        public static OcPropertyChangedListener<T> Create<T>(ObservableCollection<T> collection, string propertyName = "") where T : INotifyPropertyChanged
        {
            return new OcPropertyChangedListener<T>(collection, propertyName);
        }
    }
    #endregion

    #region TwoColumnGrid

    /// <summary>
    /// Defines a table that has two columns with any number of rows. 
    /// </summary>
    /// <remarks>
    /// This panel is designed for use in configuration/settings windows where you typically
    /// have a pairs of "Label: SomeControl" organized in rows.
    /// 
    /// The width of the first column is determined by the widest item that column and the width of the 
    /// second column is expanded to occupy all remaining space.
    /// 
    /// Written by: Isak Savo, isak.savo@gmail.com
    /// Licensed under the Code Project Open License http://www.codeproject.com/info/cpol10.aspx
    /// 
    /// XAML Example:
    /// <local:TwoColumnGrid RowSpacing="{Binding ElementName=rowSpacing, Path=Value}" ColumnSpacing="{Binding ElementName=colSpacing, Path=Value}">
    ///     <Label Content="Name:" />
    ///     <TextBox Text="John Doe" VerticalAlignment="Center" />
    ///     <Label Content="Address:" />
    ///     <TextBox Text="34 Some Street&#x0a;123 45 SomeTown&#x0a;Some Country" VerticalAlignment="Center" Height="70" />
    ///     <Label Content="Position:" />
    ///     <TextBox Text="Manager" />
    /// </local:TwoColumnGrid>
    /// 
    /// </remarks>
    public  class TwoColumnGrid : Panel
    {
        private double Column1Width;
        private List<Double> RowHeights = new List<double>();

        /// <summary>
        /// Gets or sets the amount of spacing (in device independent pixels) between the rows.
        /// </summary>
        public double RowSpacing
        {
            get { return (double)GetValue(RowSpacingProperty); }
            set { SetValue(RowSpacingProperty, value); }
        }

        /// <summary>
        /// Identifies the ColumnSpacing dependency property
        /// </summary>
        public static readonly DependencyProperty RowSpacingProperty =
            DependencyProperty.Register("RowSpacing", typeof(double), typeof(TwoColumnGrid),
            new FrameworkPropertyMetadata(0.0d, FrameworkPropertyMetadataOptions.AffectsArrange | FrameworkPropertyMetadataOptions.AffectsMeasure));

        /// <summary>
        /// Gets or sets the amount of spacing (in device independent pixels) between the columns.
        /// </summary>
        public double ColumnSpacing
        {
            get { return (double)GetValue(ColumnSpacingProperty); }
            set { SetValue(ColumnSpacingProperty, value); }
        }

        /// <summary>
        /// Identifies the ColumnSpacing dependency property
        /// </summary>
        public static readonly DependencyProperty ColumnSpacingProperty =
            DependencyProperty.Register("ColumnSpacing", typeof(double), typeof(TwoColumnGrid),
            new FrameworkPropertyMetadata(0.0d, FrameworkPropertyMetadataOptions.AffectsArrange | FrameworkPropertyMetadataOptions.AffectsMeasure));


        /// <summary>
        /// Measures the size required for all the child elements in this panel.
        /// </summary>
        /// <param name="constraint">The size constraint given by our parent.</param>
        /// <returns>The requested size for this panel including all children</returns>
        protected override Size MeasureOverride(Size constraint)
        {
            double col1Width = 0;
            double col2Width = 0;
            RowHeights.Clear();
            // First, measure all the left column children
            for (int i = 0; i < VisualChildrenCount; i += 2)
            {
                var child = Children[i];
                child.Measure(constraint);
                col1Width = Math.Max(child.DesiredSize.Width, col1Width);
                RowHeights.Add(child.DesiredSize.Height);
            }
            // Then, measure all the right column children, they get whatever remains in width
            var newWidth = Math.Max(0, constraint.Width - col1Width - ColumnSpacing);
            Size newConstraint = new Size(newWidth, constraint.Height);
            for (int i = 1; i < VisualChildrenCount; i += 2)
            {
                var child = Children[i];
                child.Measure(newConstraint);
                col2Width = Math.Max(child.DesiredSize.Width, col2Width);
                RowHeights[i / 2] = Math.Max(RowHeights[i / 2], child.DesiredSize.Height);
            }

            Column1Width = col1Width;
            return new Size(
                col1Width + ColumnSpacing + col2Width,
                RowHeights.Sum() + ((RowHeights.Count - 1) * RowSpacing));
        }

        /// <summary>
        /// Position elements and determine the final size for this panel.
        /// </summary>
        /// <param name="arrangeSize">The final area where child elements should be positioned.</param>
        /// <returns>The final size required by this panel</returns>
        protected override Size ArrangeOverride(Size arrangeSize)
        {
            double y = 0;
            for (int i = 0; i < VisualChildrenCount; i++)
            {
                var child = Children[i];
                double height = RowHeights[i / 2];
                if (i % 2 == 0)
                {
                    // Left child
                    var r = new Rect(0, y, Column1Width, height);
                    child.Arrange(r);
                }
                else
                {
                    // Right child
                    var r = new Rect(Column1Width + ColumnSpacing, y, arrangeSize.Width - Column1Width - ColumnSpacing, height);
                    child.Arrange(r);
                    y += height;
                    y += RowSpacing;
                }
            }
            return base.ArrangeOverride(arrangeSize);
        }

    }

    #endregion
    
    #region Property Converters

    #region Null2HiddenConverter, Null2CollapsedConverter

    /// <summary>
    /// Converts to Visible/Hidden.
    /// </summary>
    public class Null2HiddenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value == null ? Visibility.Hidden : Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    //
    public class False2HiddenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (bool)value ? Visibility.Visible : Visibility.Hidden;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    //
    public class True2HiddenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (bool)value ? Visibility.Hidden : Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// Converts to Visible/Collapsed.
    /// </summary>
    public class Null2CollapsedConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value == null ? Visibility.Collapsed : Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    /// <summary>
    /// Converts to Visible/Collapsed.
    /// </summary>
    public class False2CollapsedConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (bool)value ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    //
    public class True2CollapsedConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (bool)value ? Visibility.Collapsed : Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    #endregion // Null2HiddenConverter, Null2CollapsedConverter

    #endregion // Property Converters


    #region RelayCommand

    /// <summary>
    /// A command whose sole purpose is to 
    /// relay its functionality to other
    /// objects by invoking delegates. The
    /// default return value for the CanExecute
    /// method is 'true'.
    /// </summary>
    public class RelayCommand : ICommand
    {
        #region Fields

        readonly Action<object> _execute;
        readonly Predicate<object> _canExecute;

        #endregion // Fields

        #region Constructors

        /// <summary>
        /// Creates a new command that can always execute.
        /// </summary>
        /// <param name="execute">The execution logic.</param>
        public RelayCommand(Action<object> execute)
            : this(execute, null)
        {
        }

        /// <summary>
        /// Creates a new command.
        /// </summary>
        /// <param name="execute">The execution logic.</param>
        /// <param name="canExecute">The execution status logic.</param>
        public RelayCommand(Action<object> execute, Predicate<object> canExecute)
        {
            if (execute == null)
                throw new ArgumentNullException("execute");

            _execute = execute;
            _canExecute = canExecute;
        }

        #endregion // Constructors

        #region ICommand Members

        [DebuggerStepThrough]
        public bool CanExecute(object parameters)
        {
            return _canExecute == null ? true : _canExecute(parameters);
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public void Execute(object parameters)
        {
            _execute(parameters);
        }

        #endregion // ICommand Members
    }

    #endregion // RelayCommand


    public static class ColorAndBrush
    {

        public static Color MidColor(Color c1, Color c2, float percent)
        {
            // calculates a color between c1 and c2, percent being the relative position (should be 0..100)
            // percent = 0:   returns pure c1
            // percent = 100: returns pure c2

            if (percent < 0)
                percent = 0;
            else if (percent > 100)
                percent = 1;
            else
                percent /= 100;

            return c1 * (1 - percent) + c2 * percent;
        }

        public static Color MidColor(Color c1, Color c2, Color c3, float percent, float midPoint = 50)
        {
            // calculates a color between c1 and c2, percent being the relative position (should be 0..100)
            // percent = 0:   returns pure c1
            // percent = midpoint: returns pure c2
            // percent = 100: returns pure c3

            if (percent < 0)
                percent = 0;
            else if (percent > 100)
                percent = 1;
            else
                percent /= 100;

            midPoint /= 100;

            if (percent <= midPoint)
            {
                percent /= midPoint;
                return c1 * (1 - percent) + c2 * percent;
            }
            else
            {
                percent = (percent - midPoint) / (1 - midPoint);
                return c2 * (1 - percent) + c3 * percent;
            }
        }

        public static Brush MidBrush(Color c1, Color c2, double percent)
        {
            return new SolidColorBrush(MidColor(c1, c2, (float)percent));
        }

        public static Brush MidBrush(Color c1, Color c2, Color c3, double percent, double midPoint = 50)
        {
            return new SolidColorBrush(MidColor(c1, c2, c3, (float)percent, (float)midPoint));
        }

        public static Brush ToBrush(this string s)
        {
            // usage: ColorAndBrush.ToBrush("Khaki") or ColorAndBrush.ToBrush("#FFF0C0C0")
            return (SolidColorBrush)(new BrushConverter().ConvertFrom(s));
        }


    }

    #region XElemXPath


    public class XElemXPath
    {
        public XElement xElem = null;

        public XDocument xDoc
        {
            get
            {
                return new XDocument(xElem);
            }
            set
            {
                xElem = new XElement(value.Root);
            }
        }


        public XElemXPath(XDocument xDoc)
        {
            this.xDoc = xDoc;
        }

        public XElemXPath(XElement xElement)
        {
            xElem = new XElement(xElement);
        }

        public XElemXPath(string xString)
        {
            xDoc = XDocument.Parse(xString);
        }

        public XElemXPath(XElemXPath xxObj)
        {
            xElem = new XElement(xxObj.xElem);
        }


        public void RemoveNodes(string xPath)
        {
            xElem.XPathSelectElements(xPath).ToList().ForEach(_e => _e.Remove());
        }


        public string GetValue()
        {
            return xElem.Value;
        }

        public XElemXPath GetNode(string xPath)
        {
            var elem = xElem.XPathSelectElement(xPath);
            if (elem == null)
                return null;
            return new XElemXPath(elem);
        }

        public List<XElemXPath> GetNodes(string xPath)
        {
            return (from x in xElem.XPathSelectElements(xPath) select new XElemXPath(x)).ToList();
        }

        public int GetCount(string xPath)
        {
            return GetNodes(xPath).Count;
        }

        public string GetString(string xPath = ".", string defValue = null)
        {
            object result = ((IEnumerable<Object>)xElem.XPathEvaluate(xPath)).FirstOrDefault();
            if (result is XElement)
                return ((XElement)result).Value;
            else if (result is XAttribute)
                return ((XAttribute)result).Value;
            return defValue;
        }

        public List<string> GetStrings(string xPath)
        {
            return (from x in xElem.XPathSelectElements(xPath) select x.Value).ToList();
        }

        public int GetInt(string xPath = ".", int defValue = int.MinValue)
        {
            object result = ((IEnumerable<Object>)xElem.XPathEvaluate(xPath)).FirstOrDefault();
            if (result is XElement)
                return int.Parse(((XElement)result).Value);
            else if (result is XAttribute)
                return int.Parse(((XAttribute)result).Value);
            return defValue;
        }

        // static

        public static string ToLiteral(string s)
        {
            // performs the ridiculously complex task to escape an arbitrary String for XPath
            // see also: http://stackoverflow.com/questions/1341847/special-character-in-xpath-query

            if (s == null)
                throw new ArgumentException("Must not be Nothing");

            if (!s.Contains(@""""))
            {
                // contains no ", so put into ""
                return @"""" + s + @"""";
            }
            else if (!s.Contains("'"))
            {
                // contains no ', so put into ''
                return "'" + s + "'";
            }

            // this damn string contains both types of delimiters..
            Match _parts = Regex.Match(s, @"^([^']+|[^""]+)+");
            return string.Format("concat({0})", string.Join(",", (from Capture _part in _parts.Groups[1].Captures select ToLiteral(_part.Value))));
        }
    }

    #endregion  // XElemXPath


}
