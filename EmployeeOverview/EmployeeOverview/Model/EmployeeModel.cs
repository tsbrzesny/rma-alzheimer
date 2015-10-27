using System;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using RMA2_Roots;


namespace Alzheimer.Model
{
    /// <summary>
    /// 
    /// </summary>
    public class EmployeeModel
    {

        private RMA2D.RMAMitarbeiter _rmaMitarbeiter;

        public EmployeeModel(RMA2D.RMAMitarbeiter mitarbeiter)
        {
            _rmaMitarbeiter = mitarbeiter;
        }



        #region Getter and Setter
        public int ContactId
        {
            get { return _rmaMitarbeiter.contact_id; }
            set { _rmaMitarbeiter.contact_id = value; }
        }

        public string Vorname
        {
            get { return _rmaMitarbeiter.vorname; }
            set { _rmaMitarbeiter.vorname = value; }
        }

        public string Nachname
        {
            get { return _rmaMitarbeiter.nachname; }
            set { _rmaMitarbeiter.nachname = value; }
        }

        public string Zeichen
        {
            get { return _rmaMitarbeiter.zeichen; }
            set { _rmaMitarbeiter.zeichen = value; }
        }

        public string Email
        {
            get { return _rmaMitarbeiter.email; }
            set { _rmaMitarbeiter.email = value; }
        }

        public ImageSource Image
        {
            get
            {
                var fi = _rmaMitarbeiter.GetContactImagePath();
                if (fi == null)
                    return null;

                BitmapImage logo = new BitmapImage();
                logo.BeginInit();
                logo.UriSource = new Uri(fi.FullName);
                logo.EndInit();
                return logo; 
            }
        }

        public string FullName
        {
            get { return _rmaMitarbeiter.vorname + " " + _rmaMitarbeiter.nachname; }
        }
        #endregion

    }
}
