using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.Utility;


namespace ExportOBJ
{
    [Transaction(TransactionMode.Automatic)]
    class Command : IFaceEmitter, IExternalCommand
    {
        string _exportFolderName = "C:\\tmp\\export_data";
        //VertexLookupXyz _vertices;
        VertexLookupInt _vertices;

        /// <summary>
        /// List of triangles, defined as 
        /// triples of vertex indices.
        /// </summary>
        List<int> _triangles;

        /// <summary>
        /// Keep track of the number of faces processed.
        /// </summary>
        int _faceCount;

        /// <summary>
        /// Keep track of the number of triangles processed.
        /// </summary>
        int _triangleCount;

        public Command()
        {
            _faceCount = 0;
            _triangleCount = 0;
            _vertices = new VertexLookupInt();
            _triangles = new List<int>();
        }

        /// <summary>
        /// Add the vertices of the given triangle to our
        /// vertex lookup dictionary and emit a triangle.
        /// </summary>
        void StoreTriangle(MeshTriangle triangle)
        {
            for (int i = 0; i < 3; ++i)
            {
                XYZ p = triangle.get_Vertex(i);
                PointInt q = new PointInt(p);
                _triangles.Add(_vertices.AddVertex(q));
            }
        }

        /// <summary>
        /// Emit a Revit geometry Face object and 
        /// return the number of resulting triangles.
        /// </summary>
        public int EmitFace(Face face, Autodesk.Revit.DB.Color color)
        {
            ++_faceCount;
            Mesh mesh = face.Triangulate();
            int n = mesh.NumTriangles;
            Debug.Print( " {0} mesh triangles", n );
 
            for( int i = 0; i < n; ++i )
            {
                ++_triangleCount;
                MeshTriangle t = mesh.get_Triangle( i );
                StoreTriangle( t );
            }

            return n;
        }

        public int GetFaceCount()
        {
            return _faceCount;
        }

        /// <summary>
        /// Return the number of triangles processed.
        /// </summary>
        public int GetTriangleCount()
        {
            int n = _triangles.Count;

            Debug.Assert(0 == n % 3,
              "expected a multiple of 3");
            Debug.Assert(_triangleCount.Equals(n / 3),
              "expected equal triangle count");

            return _triangleCount;
        }

        public int GetVertexCount()
        {
            return _vertices.Count;
        }

        #region ExportTo: output the OBJ file
        /// <summary>
        /// Obsolete: emit an XYZ vertex.
        /// </summary>
        static void EmitVertex(StreamWriter s, XYZ p)
        {
            //s.WriteLine("v {0} {1} {2}",
            //  Util.RealString(p.X),
            //  Util.RealString(p.Y),
            //  Util.RealString(p.Z));
        }

        /// <summary>
        /// Emit a vertex to OBJ. The first vertex listed 
        /// in the file has index 1, and subsequent ones
        /// are numbered sequentially.
        /// </summary>
        static void EmitVertex(
          StreamWriter s,
          PointInt p)
        {
            s.WriteLine("v {0} {1} {2}", p.X, p.Y, p.Z);
        }

        /// <summary>
        /// Emit an OBJ triangular face.
        /// </summary>
        static void EmitFacet(
          StreamWriter s,
          int i,
          int j,
          int k)
        {
            s.WriteLine("f {0} {1} {2}",
              i + 1, j + 1, k + 1);
        }

        public void ExportTo(string path)
        {
            using (StreamWriter s = new StreamWriter(path))
            {
                foreach (PointInt key in _vertices.Keys)
                {
                    EmitVertex(s, key);
                }

                int i = 0;
                int n = _triangles.Count;

                while (i < n)
                {
                    int i1 = _triangles[i++];
                    int i2 = _triangles[i++];
                    int i3 = _triangles[i++];

                    EmitFacet(s, i1, i2, i3);
                }
            }
        }
        #endregion // ExportTo: output the OBJ file

        /// <summary>
        /// 
        /// </summary>
        /// <param name="e"></param>
        /// <param name="opt"></param>
        /// <returns></returns>
        Solid GetSolid(Element e, Options opt)
        {
            Solid solid = null;
            // access solid geometry of element
            GeometryElement geo = e.get_Geometry(opt);

            if (null != geo)
            {
                if (e is FamilyInstance)
                {
                    geo = geo.GetTransformed(Transform.Identity);
                }

                GeometryInstance inst = null;
                // iterate through multiple geometry objects
                foreach (GeometryObject obj in geo)
                {
                    solid = obj as Solid;
                    // if there is a solid with faces stop here
                    if (null != solid && 0 < solid.Faces.Size)
                    {
                        break;
                    }
                    // otherwise assign variable inst
                    inst = obj as GeometryInstance;
                }
                // if there is an instance def but no solid check the symbol
                if (null == solid && null != inst)
                {
                    geo = inst.GetSymbolGeometry();
                    // iterate through geometry of symbol
                    foreach (GeometryObject obj in geo)
                    {
                        solid = obj as Solid;
                        // if there is a solid with faces stop here
                        if (null != solid && 0 < solid.Faces.Size)
                        {
                            break;
                        }
                    }
                }
            }
            return solid;
        }

        /// <summary>
        /// Recursive method that checks for groups and gets solid geometry from elements
        /// </summary>
        /// <param name="emitter"></param>
        /// <param name="e"></param>
        /// <param name="opt"></param>
        /// <returns>int</returns>
        int ExportElement(IFaceEmitter emitter, Element e, Options opt)
        {
            Group group = e as Group;

            // if element successfully casts to a group then iterate through the group
            if (null != group)
            {
                int n = 0;

                foreach (ElementId id in group.GetMemberIds())
                {
                    Element e2 = e.Document.GetElement(id);
                    n += ExportElement(emitter, e2, opt);
                }
                return n;
            }

            // return if element has no category
            if (null == e.Category)
            {
                return 0;
            }

            // access the geometry object
            Solid solid = GetSolid(e, opt);
            // return if no solid
            if (null == solid)
            {
                return 0;
            }

            Material material;
            Color color;

            foreach (Face face in solid.Faces)
            {
                material = e.Document.GetElement(face.MaterialElementId) as Material;
                // if no material, no color
                color = (null == material) ? null : material.Color;

                emitter.EmitFace(face, color);
            }
            return 1;
        }

        /// <summary>
        /// Helper function to export elements
        /// </summary>
        /// <param name="emitter"></param>
        /// <param name="collector"></param>
        /// <param name="opt"></param>
        void ExportElements(IFaceEmitter emitter, FilteredElementCollector collector, Options opt)
        {
            int nElements = 0;
            int nSolids = 0;

            foreach (Element e in collector)
            {
                ++nElements;
                nSolids += ExportElement(emitter, e, opt);
            }

            int nFaces = emitter.GetFaceCount();
            int nTriangles = emitter.GetTriangleCount();
            int nVertices = emitter.GetVertexCount();
        }

        /// <summary>
        /// Program entry point
        /// </summary>
        /// <param name="commandData"></param>
        /// <param name="message"></param>
        /// <param name="elements"></param>
        /// <returns>Result</returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
                Document doc = uidoc.Document;

                // Determine elements to export
                FilteredElementCollector collector = null;

                // Access current selection
                SelElementSet set = uidoc.Selection.Elements;

                int n = set.Size;

                if (0 < n)
                {
                    // If any elements were preselected, export those
                    ICollection<ElementId> ids = set
                      .Cast<Element>()
                      .Select<Element, ElementId>(e => e.Id)
                      .ToArray<ElementId>();

                    collector = new FilteredElementCollector(doc, ids);
                }
                else
                {
                    // If nothing was preselected, export everything
                    collector = new FilteredElementCollector(doc);
                }

                collector.WhereElementIsNotElementType()
                    .WhereElementIsViewIndependent();

                if (null == _exportFolderName)
                {
                    _exportFolderName = Path.GetTempPath();
                }

                string filename = null;

                if (!FileSelect(_exportFolderName,
                  out filename))
                {
                    return Result.Cancelled;
                }

                _exportFolderName = Path.GetDirectoryName(filename);
                Command exporter = new Command();
                Options opt = app.Create.NewGeometryOptions();
                ExportElements(exporter, collector, opt);
                exporter.ExportTo(filename);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        static bool FileSelect(string folder, out string filename)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Title = "Save file as";
            dlg.InitialDirectory = folder;
            bool rc = (DialogResult.OK == dlg.ShowDialog());
            filename = dlg.FileName;
            return rc;
        }
    }
}
