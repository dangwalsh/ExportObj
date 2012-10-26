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
        string _export_folder_name = "C:\\tmp\\export_data";
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

        Solid GetSolid(Element e, Options opt)
        {
            Solid solid = null;

            GeometryElement geo = e.get_Geometry(opt);

            if (null != geo)
            {
                if (e is FamilyInstance)
                {
                    geo = geo.GetTransformed(
                      Transform.Identity);
                }

                GeometryInstance inst = null;
                //Transform t = Transform.Identity;

                foreach (GeometryObject obj in geo)
                {
                    solid = obj as Solid;

                    if (null != solid
                      && 0 < solid.Faces.Size)
                    {
                        break;
                    }

                    inst = obj as GeometryInstance;
                }

                if (null == solid && null != inst)
                {
                    geo = inst.GetSymbolGeometry();
                    //t = inst.Transform;

                    foreach (GeometryObject obj in geo)
                    {
                        solid = obj as Solid;

                        if (null != solid
                          && 0 < solid.Faces.Size)
                        {
                            break;
                        }
                    }
                }
            }
            return solid;
        }

        int ExportElement(
            IFaceEmitter emitter,
            Element e,
            Options opt)
        {
            Group group = e as Group;

            if (null != group)
            {
                int n = 0;

                foreach (ElementId id
                  in group.GetMemberIds())
                {
                    Element e2 = e.Document.GetElement(
                      id);

                    n += ExportElement(emitter, e2, opt);
                }
                return n;
            }

           //string desc = Util.ElementDescription(e);

            if (null == e.Category)
            {
                //Debug.Print("Element '{0}' has no "
                //  + "category.", desc);

                return 0;
            }

            Solid solid = GetSolid(e, opt);

            if (null == solid)
            {
                //Debug.Print("Unable to access "
                //  + "solid for element {0}.", desc);

                return 0;
            }

            Material material;
            Color color;

            foreach (Face face in solid.Faces)
            {
                material = e.Document.GetElement(
                  face.MaterialElementId) as Material;

                color = (null == material)
                  ? null
                  : material.Color;

                emitter.EmitFace(face, color);
            }
            return 1;
        }

        void ExportElements(
            IFaceEmitter emitter,
            FilteredElementCollector collector,
            Options opt)
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

            //string msg = string.Format(
            //  "{0} element{1} with {2} solid{3}, "
            //  + "{4} face{5}, {6} triangle{7} and "
            //  + "{8} vertice{9} exported.",
            //  nElements, Util.PluralSuffix(nElements),
            //  nSolids, Util.PluralSuffix(nSolids),
            //  nFaces, Util.PluralSuffix(nFaces),
            //  nTriangles, Util.PluralSuffix(nTriangles),
            //  nVertices, Util.PluralSuffix(nVertices));

            //InfoMsg(msg);
        }

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
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
                    // If any elements were preselected,
                    // export those to OBJ

                    ICollection<ElementId> ids = set
                      .Cast<Element>()
                      .Select<Element, ElementId>(e => e.Id)
                      .ToArray<ElementId>();

                    collector = new FilteredElementCollector(doc, ids);
                }
                else
                {
                    // If nothing was preselected, export 
                    // all model elements to OBJ

                    collector = new FilteredElementCollector(doc);
                }

                collector.WhereElementIsNotElementType()
                    .WhereElementIsViewIndependent();

                if (null == _export_folder_name)
                {
                    _export_folder_name = Path.GetTempPath();
                }

                string filename = null;

                if (!FileSelect(_export_folder_name,
                  out filename))
                {
                    return Result.Cancelled;
                }

                _export_folder_name
                  = Path.GetDirectoryName(filename);

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

        static bool FileSelect(
            string folder,
            out string filename)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Title = "Save file as";
            //dlg.CheckFileExists = true;
            //dlg.CheckPathExists = true;
            //dlg.RestoreDirectory = true;
            dlg.InitialDirectory = folder;
            //dlg.Filter = ".txt Files (*.txt)|*.txt";
            bool rc = (DialogResult.OK == dlg.ShowDialog());
            filename = dlg.FileName;
            return rc;
        }
    }
}
