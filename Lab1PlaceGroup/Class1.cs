using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace Lab1PlaceGroup
{
    [Transaction(TransactionMode.ReadOnly)]
    [Regeneration(RegenerationOption.Manual)]
    public class Class1 : IExternalCommand
    {
        public class MySelectionFilter : ISelectionFilter
        {
            Document m_doc = null;

            public bool AllowElement(Element element)
            {
                return element is Wall || element is Duct;
            }
            public bool AllowReference(Reference refer, XYZ point)
            {
                GeometryObject geoObject =
                m_doc.GetElement(refer)
                     .GetGeometryObjectFromReference(refer);
                return geoObject != null && geoObject is Face;
            }
        }

        public Result Execute(ExternalCommandData commandData,
          ref string message, ElementSet elements)
        {
            ////Get application and document objects
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            try
            {
                while (true)
                {
                    Reference selRefWall =
                      uidoc.Selection.PickObject(ObjectType.Element,
                        new MySelectionFilter(), "select a wall");

                    Wall wall = uidoc.Document.GetElement(selRefWall) as Wall;

                    Reference selRefDuct =
                      uidoc.Selection.PickObject(ObjectType.Element,
                        new MySelectionFilter(), "select a duct");

                    Duct duct = uidoc.Document.GetElement(selRefDuct) as Duct;

                    GeometryElement geomWall = wall.get_Geometry(new Options());
                    GeometryElement geomDuct = duct.get_Geometry(new Options());

                    List<Solid> wall_solids = new List<Solid>();

                    List<Curve> wall_geo = new List<Curve>();

                    foreach (GeometryObject geomObj in geomWall)
                    {
                        Curve geomCurve = geomObj as Curve;
                        if (null != geomCurve)
                        {
                            wall_geo.Add(geomCurve);
                        }

                        Solid geomSolid = geomObj as Solid;
                        if (null != geomSolid)
                        {
                            wall_solids.Add(geomSolid);
                        }
                    }

                    List<Solid> duct_solids = new List<Solid>();

                    List<Curve> duct_geo = new List<Curve>();

                    foreach (GeometryObject geomObj in geomDuct)
                    {
                        Curve geomCurve = geomObj as Curve;
                        if (null != geomCurve)
                        {
                            duct_geo.Add(geomCurve);
                        }

                        Solid geomSolid = geomObj as Solid;
                        if (null != geomSolid)
                        {
                            duct_solids.Add(geomSolid);
                        }
                    }

                    Solid intersection = BooleanOperationsUtils.ExecuteBooleanOperation(wall_solids[0], duct_solids[0], BooleanOperationsType.Intersect);

                    TaskDialog.Show("Intersection Volume", intersection.Volume.ToString());
                }

            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException) { }

            return Result.Succeeded;
        }
    }
}
