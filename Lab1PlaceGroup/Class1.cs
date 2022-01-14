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
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            //Sudedami visi modelio elementai "wall" į sąrašą
            FilteredElementCollector WallCollector = new FilteredElementCollector(doc);
            WallCollector.OfClass(typeof(Wall));
            List<Wall> walls = WallCollector.Cast<Wall>().ToList();

            //Sudedami visi modelio elementai "duct" į sąrašą
            FilteredElementCollector DuctCollector = new FilteredElementCollector(doc);
            DuctCollector.OfClass(typeof(Duct));
            List<Duct> ducts = DuctCollector.Cast<Duct>().ToList();

            foreach (Duct du in ducts)
            {
                foreach (Wall wa in walls)
                {
                    //Surandama duct centro linija
                    Curve ductCurve = FindDuctCurve(du);

                    //Surandamas wall paviršius
                    Face wallFace = FindWallFace(wa);

                    if (wallFace != null)
                    {
                        //Ieškomas wall paviršiaus ir duct linijos susikirtimo taškas
                        XYZ intersectPoint = FindIntersectPoint(wallFace, ductCurve);

                        if (!intersectPoint.IsAlmostEqualTo(new XYZ(0, 0, 0)))
                        {
                            //Surandamas reikiamo elemento simbolis
                            FamilySymbol symbol = GetFamilySymbolByName(doc, "M_Round Face Opening");

                            //Šiuo atveju turėtų būti surandama elemento kryptis
                            XYZ refDir1 = intersectPoint.CrossProduct(XYZ.BasisZ);

                            using (Transaction transaction = new Transaction(doc))
                            {
                                transaction.Start("LetsDoThis");

                                //Iterpiamas elementas apskaičiuotame wall ir duct susikirtimo taške
                                FamilyInstance instance = doc.Create.NewFamilyInstance(wallFace, intersectPoint, refDir1, symbol);

                                transaction.Commit();
                            }
                        }
                    }
                }
            }

            return Result.Succeeded;
        }

        public static FamilySymbol GetFamilySymbolByName(Document doc, string name)
        {
            var paramId = new ElementId(BuiltInParameter.ALL_MODEL_FAMILY_NAME);
            var paramValueProvider = new ParameterValueProvider(paramId);
            var equalsRule = new FilterStringEquals();
            var filterRule = new FilterStringRule(paramValueProvider, equalsRule, name, false);
            var filter = new ElementParameterFilter(filterRule);

            var fec = new FilteredElementCollector(doc);
            fec.OfClass(typeof(FamilySymbol)).WhereElementIsElementType().WherePasses(filter);

            if (fec.GetElementCount() == 1)
            {
                var symbol = fec.FirstElement() as FamilySymbol;
                if (!symbol.IsActive)
                {
                    symbol.Activate();
                    doc.Regenerate();
                }
                return symbol;
            }
            return null;
        }

        public static Curve FindDuctCurve(Duct duct)
        {
            LocationCurve positionCurve = duct.Location as LocationCurve;

            Curve curve = positionCurve.Curve;

            return curve;
        }

        public static Face FindWallFace(Wall wall)
        {
            Face face = null;
            Options geomOptions = new Options();
            geomOptions.ComputeReferences = true;

            GeometryElement wallGeom = wall.get_Geometry(geomOptions);
            foreach (GeometryObject geomObj in wallGeom)
            {
                Solid geomSolid = geomObj as Solid;
                if (null != geomSolid)
                {
                    foreach (Face geomFace in geomSolid.Faces)
                    {
                        face = geomFace;
                        break;
                    }
                    break;
                }
            }

            return face;
        }

        public static XYZ FindIntersectPoint(Face wallFace, Curve ductCirve)
        {
            IntersectionResultArray results;
            var inter = wallFace.Intersect(ductCirve, out results);
            XYZ point = new XYZ(0,0,0);

            if (inter.ToString().Contains("Overlap"))
            {
                IntersectionResult iResult = results.get_Item(0);
                point = iResult.XYZPoint;
            }

            return point;
        }
    }
}
