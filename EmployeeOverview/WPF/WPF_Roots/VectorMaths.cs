using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Windows;
using System.Windows.Media.Imaging;


namespace WPF_Roots
{


    public class Vector2D
    {
        public float x, y = 0;

        public Vector2D(float x, float y)
        {
            this.x = x;
            this.y = y;
        }

        public Vector2D(PointF point)
        {
            this.x = point.X;
            this.y = point.Y;
        }

        public Vector2D(SizeF size)
        {
            this.x = size.Width;
            this.y = size.Height;
        }

        public Vector2D(System.Windows.Point point)
        {
            this.x = (float)point.X;
            this.y = (float)point.Y;
        }

        public float this[int index]
        {
            get
            {
                switch (index) {
                    case 0:
                        return x;
                    case 1:
                        return y;
                    default:
                        throw new ArgumentNullException();
                }
            }
            set
            {
                switch (index) {
                    case 0:
                        x = value;
                        break;
                    case 1:
                        y = value;
                        break;
                    default:
                        throw new ArgumentNullException();
                }
            }
        }



        // operators

        public static Vector2D operator -(Vector2D v)
        {
            return new Vector2D(0,0).Subtract(v);
        }

        public static Vector2D operator +(Vector2D v1, Vector2D v2)
        {
            return v1.Add(v2);
        }

        public static PointF operator +(PointF p, Vector2D v)
        {
            return new PointF(p.X + v.x, p.Y + v.y);
        }

        public static Vector2D operator -(Vector2D v1, Vector2D v2)
        {
            return v1.Subtract(v2);
        }

        public static PointF operator -(PointF p, Vector2D v)
        {
            return new PointF(p.X - v.x, p.Y - v.y);
        }

        public static Vector2D operator *(Vector2D v, float by)
        {
            return v.Multiply(by);
        }

        public static Vector2D operator *(float by, Vector2D v)
        {
            return v.Multiply(by);
        }


        // members

        // returns (this + v)
        public Vector2D Add(Vector2D v)
        {
            float _x = x + v.x;
            float _y = y + v.y;

            return new Vector2D(_x, _y);
        }

        // returns (this - v)
        public Vector2D Subtract(Vector2D v)
        {
            float _x = x - v.x;
            float _y = y - v.y;

            return new Vector2D(_x, _y);
        }

        // returns (this x v)
        public Vector3D Cross(Vector2D with)
        {
            float _z = x * with.y - y * with.x;
            return new Vector3D(0, 0, _z);
        }

        // returns the normalized unit vector of this
        public Vector2D Normalize()
        {
            float l = this.Length;
            float _x = x / l;
            float _y = y / l;
            return new Vector2D(_x, _y);
        }

        // returns a scaled version of (this)
        public Vector2D Multiply(float by)
        {
            float _x = x * by;
            float _y = y * by;
            return new Vector2D(_x, _y);
        }

        // returns a scaled version of (this)
        public Vector2D Multiply(float by_x, float by_y)
        {
            float _x = x * by_x;
            float _y = y * by_y;
            return new Vector2D(_x, _y);
        }

        // returns the dot product (this . with)
        public float Dot(Vector2D with)
        {
            return x * with.x + y * with.y;
        }

        // returns the length of (this)
        public float Length
        {
            get
            {
                return (float)Math.Sqrt(LengthSquared);
            }
        }

        // returns the squared length of (this)
        public float LengthSquared
        {
            get
            {
                return Dot(this);
            }
        }

        // returns true, if (this) points into the given rectangle
        public bool IsIn(Vector2D r1, Vector2D r2)
        {
            float _x = Math.Min(r1.x, r2.x);
            float _xs = Math.Abs(r1.x - r2.x);
            float _y = Math.Min(r1.y, r2.y);
            float _ys = Math.Abs(r1.y - r2.y);

            if (y < _y || y > _y + _ys)
                return false;
            if (x < _x || x > _x + _xs)
                return false;

            return true;
        }

        public bool IsIn(RectangleF rect)
        {
            if (y < rect.Top || y > rect.Bottom)
                return false;
            if (x < rect.Left || x > rect.Right)
                return false;

            return true;
        }



        // returns the minimum by components of this and the given vector
        public Vector2D MinByComponent(Vector2D v)
        {
            return new Vector2D(Math.Min(x, v.x), Math.Min(y, v.y));
        }

        // returns the maximum by components of this and the given vector
        public Vector2D MaxByComponent(Vector2D v)
        {
            return new Vector2D(Math.Max(x, v.x), Math.Max(y, v.y));
        }

        //
        public override String ToString()
        {
            return x + "," + y;
        }

        //
        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType())
                return false;

            var otherV = (Vector2D)obj;
            return (otherV.x == x && otherV.y == y);
        }

        public override int GetHashCode()
        {
            return x.GetHashCode() ^ y.GetHashCode();
        }


        // transformations to other types

        public PointF ToPointF()
        {
            return new PointF(x, y);
        }

        public SizeF ToSizeF()
        {
            return new SizeF(Math.Abs(x), Math.Abs(y));
        }

        // two Vector2D to Rectangle
        public static RectangleF Rect(Vector2D v1, Vector2D v2)
        {
            // returns the Rectangle defined by the diagonal points v1 and v2
            float x = Math.Min(v1.x, v2.x);
            float sx = Math.Abs(v1.x - v2.x);
            float y = Math.Min(v1.y, v2.y);
            float sy = Math.Abs(v1.y - v2.y);
            return new RectangleF(x, y, sx, sy);
        }

        // two PointF to Rectangle
        public static RectangleF Rect(PointF p1, PointF p2)
        {
            // returns the Rectangle defined by the diagonal points p1 and p2
            float x = Math.Min(p1.X, p2.X);
            float sx = Math.Abs(p1.X - p2.X);
            float y = Math.Min(p1.Y, p2.Y);
            float sy = Math.Abs(p1.Y - p2.Y);
            return new RectangleF(x, y, sx, sy);
        }

        public Vector3D ToVector3D()
        {
            return new Vector3D(x, y, 0);
        }

    }


    public static class Extensions_Rotation
    {
        public static Rotation ReverseRotation(this Rotation rotation)
        {
            switch (rotation)
            {
                case Rotation.Rotate90:
                    return Rotation.Rotate270;
                case Rotation.Rotate270:
                    return Rotation.Rotate90;
                default:
                    break;
            }
            return rotation;
        }

        public static Rotation Add90(this Rotation rotation)
        {
            switch (rotation)
            {
                case Rotation.Rotate0:
                    return Rotation.Rotate90;
                case Rotation.Rotate90:
                    return Rotation.Rotate180;
                case Rotation.Rotate180:
                    return Rotation.Rotate270;
                default:
                    break;
            }
            return Rotation.Rotate0;
        }

        public static Rotation Add180(this Rotation rotation)
        {
            switch (rotation)
            {
                case Rotation.Rotate0:
                    return Rotation.Rotate180;
                case Rotation.Rotate90:
                    return Rotation.Rotate270;
                case Rotation.Rotate270:
                    return Rotation.Rotate90;
                default:
                    break;
            }
            return Rotation.Rotate0;
        }

        public static Rotation Sub90(this Rotation rotation)
        {
            switch (rotation)
            {
                case Rotation.Rotate0:
                    return Rotation.Rotate270;
                case Rotation.Rotate270:
                    return Rotation.Rotate180;
                case Rotation.Rotate180:
                    return Rotation.Rotate90;
                default:
                    break;
            }
            return Rotation.Rotate0;
        }

        public static bool IsZero(this Rotation rotation)
        {
            return (rotation == Rotation.Rotate0);
        }
        public static int ToDegrees(this Rotation rotation)
        {
            switch (rotation)
            {
                case Rotation.Rotate90:
                    return 90;
                case Rotation.Rotate180:
                    return 180;
                case Rotation.Rotate270:
                    return 270;
                default:
                    break;
            }
            return 0;
        }
    }


    public static class Extensions_RectangleF
    {
        public static PointF Center(this RectangleF r)
        {
            return new PointF((r.Left + r.Right) / 2, (r.Bottom + r.Top) / 2);
        }

        public static Vector2D BR_TL(this RectangleF r)
        {
            return new Vector2D(r.Right - r.Left, r.Bottom - r.Top);
        }
    }


    public class Matrix2x2
    {
        public float x1y1, x2y1 = 0.0f;
        public float x1y2, x2y2 = 0.0f;

        public Matrix2x2(float x1y1, float x2y1, float x1y2, float x2y2)
        {
            this.x1y1 = x1y1;
            this.x2y1 = x2y1;
            this.x1y2 = x1y2;
            this.x2y2 = x2y2;
        }

        public static Matrix2x2 Unit(float unit = 1f)
        {
            return new Matrix2x2(unit, 0f, 0f, unit);
        }

        public static Matrix2x2 RotationMatrix(Rotation rotation, float unit = 1f)
        {
            switch (rotation)
            {
                case Rotation.Rotate90:
                    return new Matrix2x2(0f, -unit, unit, 0f);
                case Rotation.Rotate180:
                    return new Matrix2x2(-unit, 0f, 0f, -unit);
                case Rotation.Rotate270:
                    return new Matrix2x2(0f, unit, -unit, 0f);
                default:
                    break;
            }
            return Matrix2x2.Unit(unit);
        }

        public static Matrix2x2 ReverseRotationMatrix(Rotation rotation, float unit = 1f)
        {
            return RotationMatrix(rotation.ReverseRotation(), unit);
        }

        public static Vector2D operator *(Matrix2x2 m, Vector2D v)
        {
            return new Vector2D(m.x1y1 * v.x + m.x2y1 * v.y, m.x1y2 * v.x + m.x2y2 * v.y);
        }

        public static SizeF operator *(Matrix2x2 m, SizeF s)
        {
            return new SizeF(Math.Abs(m.x1y1 * s.Width + m.x2y1 * s.Height), Math.Abs(m.x1y2 * s.Width + m.x2y2 * s.Height));
        }

        public static Matrix2x2 operator *(Matrix2x2 m1, Matrix2x2 m2)
        {
            return new Matrix2x2(m1.x1y1 * m2.x1y1 + m1.x2y1 * m2.x1y2,
                                 m1.x1y1 * m2.x2y1 + m1.x2y1 * m2.x2y2,
                                 m1.x1y2 * m2.x1y1 + m1.x2y2 * m2.x1y2,
                                 m1.x1y2 * m2.x2y1 + m1.x2y2 * m2.x2y2);
        }

        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType())
                return false;

            var m = (Matrix2x2)obj;
            return (m.x1y1 == x1y1 && m.x1y2 == x1y2 && m.x2y1 == x2y1 && m.x2y2 == x2y2);
        }

        public override int GetHashCode()
        {
            return x1y1.GetHashCode() ^ x1y2.GetHashCode() ^ x2y1.GetHashCode() ^ x2y2.GetHashCode();
        }
    }


    public class Matrix3x3
    {
        public float x1y1, x2y1, x3y1 = 0.0f;
        public float x1y2, x2y2, x3y2 = 0.0f;
        public float x1y3, x2y3, x3y3 = 0.0f;

        public Matrix3x3(float x1y1, float x2y1, float x3y1,
                         float x1y2, float x2y2, float x3y2,
                         float x1y3, float x2y3, float x3y3)
        {
            this.x1y1 = x1y1;
            this.x1y2 = x1y2;
            this.x1y3 = x1y3;
            this.x2y1 = x2y1;
            this.x2y2 = x2y2;
            this.x2y3 = x2y3;
            this.x3y1 = x3y1;
            this.x3y2 = x3y2;
            this.x3y3 = x3y3;
        }

        public Matrix3x3(Matrix2x2 m2x2)
        {
            x1y1 = m2x2.x1y1;
            x1y2 = m2x2.x1y2;
            x2y1 = m2x2.x2y1;
            x2y2 = m2x2.x2y2;
        }

        public static Vector2D operator *(Matrix3x3 m, Vector2D v)
        {
            // for the product, v.z is set to 1
            // the result.z is dropped
            return new Vector2D(m.x1y1 * v.x + m.x2y1 * v.y + m.x3y1,
                                m.x1y2 * v.x + m.x2y2 * v.y + m.x3y2);
        }

        public Matrix2x2 ToMatrix2x2()
        {
            return new Matrix2x2(x1y1, x2y1, x1y2, x2y2);
        }

        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType())
                return false;

            var m = (Matrix3x3)obj;
            return (m.x1y1 == x1y1 && m.x2y1 == x2y1 && m.x3y1 == x3y1 &&
                    m.x1y2 == x1y2 && m.x2y2 == x2y2 && m.x3y2 == x3y2 &&
                    m.x1y3 == x1y3 && m.x2y3 == x2y3 && m.x3y3 == x3y3);
        }

        public override int GetHashCode()
        {
            return x1y1.GetHashCode() ^ x2y1.GetHashCode() ^ x3y1.GetHashCode() ^
                   x1y2.GetHashCode() ^ x2y2.GetHashCode() ^ x3y2.GetHashCode() ^
                   x1y3.GetHashCode() ^ x2y3.GetHashCode() ^ x3y3.GetHashCode();
        }
    }


    public class Vector3D
    {
        public float x, y, z = 0;

        public Vector3D(float x, float y, float z)
        {
            this.x = x;
            this.y = y;
            this.z = z;
        }

        public float this[int index]
        {
            get
            {
                switch (index) {
                    case 0:
                        return x;
                    case 1:
                        return y;
                    case 2:
                        return z;
                    default:
                        throw new ArgumentNullException();
                }
            }
            set
            {
                switch (index) {
                    case 0:
                        x = value;
                        break;
                    case 1:
                        y = value;
                        break;
                    case 2:
                        z = value;
                        break;
                    default:
                        throw new ArgumentNullException();
                }
            }
        }

        public Vector2D ToVector2D()
        {
            // simply gets rid of the z-component
            return new Vector2D(x, y);
        }
    }


    // Transformation: Stretch&Move, no Rotation
    public class T_StretchMove2D
    {
        private float mx, my;
        private Vector2D t;


        public T_StretchMove2D(float mx, float my, float tx, float ty)
        {
            this.mx = mx;
            this.my = my;
            this.t = new Vector2D(tx, ty);
        }


        // operators

        public static Vector2D operator *(Vector2D v, T_StretchMove2D tsm)
        {
            return tsm.ApplyTo(v);
        }

        public static Vector2D operator *(T_StretchMove2D tsm, Vector2D v)
        {
            return tsm.ApplyTo(v);
        }

        public static RectangleF operator *(RectangleF r, T_StretchMove2D tsm)
        {
            Vector2D p1 = new Vector2D((float)r.Right, (float)r.Top) * tsm;
            Vector2D p2 = new Vector2D((float)r.Left, (float)r.Bottom) * tsm;
            return Vector2D.Rect(p1, p2);
        }


        // members

        public Vector2D ApplyTo(Vector2D v)
        {
            return new Vector2D(v.x * mx, v.y * my) + t;
        }

        public T_StretchMove2D Invert()
        {
            // crashes for degenerated tsms with mx/my == 0
            return new T_StretchMove2D(1 / mx, 1 / my, -t.x / mx, -t.y / my);
        }
    }


    // Transformation: Origin --> viewport (scale, rotate, reflect, translate)
    public class T_Viewport
    {

        public static T_Viewport GetSelfTransformationFromRotation(SizeF sourceSize, Rotation targetRotation)
        {
            var unitVector = new Vector2D(1, 1);
            var targetSize = Matrix2x2.RotationMatrix(targetRotation) * sourceSize;
            return new T_Viewport(sourceSize, unitVector, targetSize, unitVector, targetRotation);
        }

        public SizeF sourceSize { get; private set; }       // size of the source rectangle
        public Vector2D sourceUnits { get; private set; }   // direction of the positive source x & y coordinates
        public SizeF targetSize { get; private set; }       // size of the target rectangle
        public Vector2D targetUnits { get; private set; }   // direction of the positive target x & y coordinates
        public Rotation targetRotation { get; private set; }

        private Matrix3x3 m;

        public T_Viewport(SizeF sourceSize, Vector2D sourceUnits,
                          SizeF targetSize, Vector2D targetUnits,
                          Rotation targetRotation)
        {
            this.sourceSize = sourceSize;
            this.sourceUnits = sourceUnits;
            this.targetSize = targetSize;
            this.targetUnits = targetUnits;
            this.targetRotation = targetRotation;

            // create 2D scaling & rotation & axis reflection
            var normalizedTargetSize = Matrix2x2.RotationMatrix(targetRotation.ReverseRotation()) * targetSize;
            float mx = normalizedTargetSize.Width / sourceSize.Width;
            float my = normalizedTargetSize.Height / sourceSize.Height;
            var reflection = Matrix2x2.Unit();
            if (targetUnits.x * sourceUnits.x < 0)
                reflection.x1y1 = -1;
            if (targetUnits.y * sourceUnits.y < 0)
                reflection.x2y2 = -1;
            Matrix2x2 scaleRotateReflect = new Matrix2x2(mx, 0, 0, my) * Matrix2x2.RotationMatrix(targetRotation) * reflection;

            // combine with translation into a 3D matrix
            m = new Matrix3x3(scaleRotateReflect);
            if (m.x1y1 < 0)
                m.x3y1 = normalizedTargetSize.Width;
            if (m.x1y2 < 0)
                m.x3y2 = normalizedTargetSize.Width;
            if (m.x2y1 < 0)
                m.x3y1 = normalizedTargetSize.Height;
            if (m.x2y2 < 0)
                m.x3y2 = normalizedTargetSize.Height;

        }


        public SizeF GetSourcePageSize(bool includingRotation = true)
        {
            if (includingRotation)
                return Matrix2x2.RotationMatrix(targetRotation) * sourceSize;
            return sourceSize;
        }


        // operators

        public static PointF operator *(T_Viewport tvp, PointF p)
        {
            return tvp.ApplyTo(p);
        }

        public static PointF operator *(T_Viewport tvp, System.Windows.Point p)
        {
            return tvp.ApplyTo(p);
        }

        public static SizeF operator *(T_Viewport tvp, SizeF s)
        {
            return tvp.m.ToMatrix2x2() * s;
        }

        public static RectangleF operator *(T_Viewport tvp, RectangleF r)
        {
            PointF p1 = tvp.ApplyTo(new PointF(r.Right, r.Top));
            PointF p2 = tvp.ApplyTo(new PointF(r.Left, r.Bottom));
            return Vector2D.Rect(p1, p2);
        }


        // members

        public PointF ApplyTo(PointF p)
        {
            var v = m * new Vector2D(p);
            return new System.Drawing.PointF(v.x, v.y);
        }

        public PointF ApplyTo(System.Windows.Point p)
        {
            var v = m * new Vector2D(p);
            return new System.Drawing.PointF(v.x, v.y);
        }

        public Vector2D ApplyToVector(Vector2D v)
        {
            // v is interpreted as a real vector - no translation applies
            return m.ToMatrix2x2() * v;
        }

        public SizeF Scale(SizeF s)
        {
            // applies the scaling factors, but no rotation/translation
            return new SizeF(Math.Abs(m.x1y1 + m.x1y2) * s.Width,
                             Math.Abs(m.x2y1 + m.x2y2) * s.Height);
        }

        public SizeF Rotate(SizeF s)
        {
            // applies the rotation factors, but no scaling/translation
            return Matrix2x2.RotationMatrix(targetRotation) * s;
        }

        public T_Viewport GetFinalSourceTransformation()
        {
            // gets the transformation that is required to find the position of objects on the rotated page
            return new T_Viewport(sourceSize, sourceUnits, Rotate(sourceSize), sourceUnits, targetRotation.ReverseRotation());
        }


        public T_Viewport Inversion
        {
            get
            {
                return new T_Viewport(targetSize, targetUnits,
                                      sourceSize, sourceUnits,
                                      targetRotation);
            }
        }

        public override bool Equals(object obj)
        {
            if (obj == null || GetType() != obj.GetType())
                return false;

            var otherT = (T_Viewport)obj;
            return (otherT.m == m);
        }

        public override int GetHashCode()
        {
            return m.GetHashCode();
        }
    }
}





