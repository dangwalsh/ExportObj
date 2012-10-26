using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ExportOBJ
{
    class VertexLookupXYZ : Dictionary<XYZ, int>
    {
        #region XyzEqualityComparer
        /// <summary>
        /// Define equality for Revit XYZ points.
        /// Very rough tolerance, as used by Revit itself.
        /// </summary>
        class XyzEqualityComparer : IEqualityComparer<XYZ>
        {
            const double _sixteenthInchInFeet
                = 1.0 / (16.0 * 12.0);

            public bool Equals(XYZ p, XYZ q)
            {
                return p.IsAlmostEqualTo(q,
                    _sixteenthInchInFeet);
            }

            public int GetHashCode(XYZ p)
            {
                //return Util.PointString(p).GetHashCode();
                return p.GetHashCode();
            }
        }
        #endregion // XyzEqualityComparer
 
        public VertexLookupXYZ()
            : base( new XyzEqualityComparer() )
        {
        }

        /// <summary>
        /// Return the index of the given vertex,
        /// adding a new entry if required.
        /// </summary>
        public int AddVertex( XYZ p )
        {
            return ContainsKey( p )
                ? this[p]
                : this[p] = Count;
        }
    }
}
