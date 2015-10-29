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

        private readonly RMA2D.RMAMitarbeiter _rmaMitarbeiter;
        private int _posX;
        private int _posY;

        public EmployeeModel(RMA2D.RMAMitarbeiter mitarbeiter, int posX = 0, int posY = 0)
        {
            _rmaMitarbeiter = mitarbeiter;
            _posX = posX;
            _posY = posY;
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

        public int PosX
        {
            get { return _posX; }
            set { _posX = value; }
        }

        public int PosY
        {
            get { return _posY;}
            set { _posY = value; }
        }

        public string FullName
        {
            get { return _rmaMitarbeiter.vorname + " " + _rmaMitarbeiter.nachname; }
        }
        #endregion

    }
}
