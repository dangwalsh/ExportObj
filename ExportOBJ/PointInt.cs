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

        const double _feetToMm = 25.4 * 12;

        /// <summary>
        /// Helper method for metric conversion
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
        static int ConvertFeetToMillimetres(double d)
        {
            return (int)(_feetToMm * d + 0.5);
        }

        /// <summary>
        /// Class constructor
        /// </summary>
        /// <param name="p"></param>
        public PointInt(XYZ p)
        {
            X = ConvertFeetToMillimetres(p.X);
            Y = ConvertFeetToMillimetres(p.Y);
            Z = ConvertFeetToMillimetres(p.Z);
        }

        /// <summary>
        /// Checks a point for colocation
        /// </summary>
        /// <param name="a"></param>
        /// <returns></returns>
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
