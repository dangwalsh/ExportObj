using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Revit.DB;

namespace ExportOBJ
{
    class PointInt : IComparable<PointInt>
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Z { get; set; }

        const double _feet_to_mm = 25.4 * 12;

        static int ConvertFeetToMillimetres(double d)
        {
            return (int)(_feet_to_mm * d + 0.5);
        }

        public PointInt(XYZ p)
        {
            X = ConvertFeetToMillimetres(p.X);
            Y = ConvertFeetToMillimetres(p.Y);
            Z = ConvertFeetToMillimetres(p.Z);
        }

        public int CompareTo(PointInt a)
        {
            int d = X - a.X;

            if (0 == d)
            {
                d = Y - a.Y;

                if (0 == d)
                {
                    d = Z - a.Z;
                }
            }
            return d;
        }
    }
}
