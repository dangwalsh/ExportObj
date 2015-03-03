using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.Revit.DB;

namespace ExportOBJ
{
    interface IFaceEmitter
    {
        /// <summary>
        /// Emit a face with a specified colour.
        /// </summary>
        int EmitFace(Face face, Color color, Transform transform);

        /// <summary>
        /// Return the final triangle count 
        /// after processing all faces.
        /// </summary>
        int GetFaceCount();

        /// <summary>
        /// Return the final triangle count 
        /// after processing all faces.
        /// </summary>
        int GetTriangleCount();

        /// <summary>
        /// Return the final vertex count 
        /// after processing all faces.
        /// </summary>
        int GetVertexCount();
    }
}
