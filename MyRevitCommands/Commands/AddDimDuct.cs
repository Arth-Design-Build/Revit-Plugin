using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class AddDimDuct : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the current document and UI application
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Get the active view
            View activeView = doc.ActiveView;

            // Get the selected elements (ducts)
            foreach (ElementId elemId in uiDoc.Selection.GetElementIds())
            {
                Element selectedElem = doc.GetElement(elemId);

                // Check if the element is a duct
                if (selectedElem is Duct duct)
                {
                    // Get the start and end points of the duct
                    XYZ startPoint = null;
                    XYZ endPoint = null;

                    // Get the geometry of the duct
                    GeometryElement geomElem = duct.get_Geometry(new Options());

                    // Iterate through the geometry instances
                    foreach (GeometryInstance geomInst in geomElem)
                    {
                        GeometryElement instGeomElem = geomInst.GetInstanceGeometry();
                        foreach (GeometryObject geomObj in instGeomElem)
                        {
                            if (geomObj is Line line)
                            {
                                // Check if the line intersects with the duct
                                IntersectionResultArray results = null;
                                Solid ductSolid = null;
                                if (geomObj is Solid solid)
                                {
                                    ductSolid = solid;
                                }
                                else if (geomObj is GeometryInstance subInst)
                                {
                                    // Iterate through the geometry objects in the subinstance to find a solid
                                    foreach (GeometryObject subGeomObj in subInst.SymbolGeometry)
                                    {
                                        if (subGeomObj is Solid subSolid)
                                        {
                                            ductSolid = subSolid;
                                            break;
                                        }
                                    }
                                }
                                if (ductSolid != null)
                                {
                                    CurveArray curveArray = new CurveArray();
                                    curveArray.Append(line);

                                    CurveLoop curveLoop = CurveLoop.Create((IList<Curve>)curveArray);

                                    Solid lineSolid = GeometryCreationUtilities.CreateExtrusionGeometry(new List<CurveLoop> { curveLoop }, XYZ.BasisZ, 0.1);

                                    // Find the intersection between the duct solid and the line solid
                                    var results1 = BooleanOperationsUtils.ExecuteBooleanOperation(ductSolid, lineSolid, BooleanOperationsType.Intersect);

                                    if (results1 != null)
                                    {
                                        if (startPoint == null)
                                        {
                                            startPoint = line.GetEndPoint(0);
                                        }
                                        else
                                        {
                                            endPoint = line.GetEndPoint(1);
                                        }
                                    }
                                }

                            }
                        }
                    }

                    // Create a dimension between the start and end points of the duct
                    if (startPoint != null && endPoint != null)
                    {
                        Line dimLine = Line.CreateBound(startPoint, endPoint);
                        Dimension dim = doc.Create.NewDimension(activeView, dimLine, null, null);

                        // Do something with the dimension
                        TaskDialog.Show("Dimension Created", $"Dimension ID: {dim.Id}");
                    }
                }
            }

            return Result.Succeeded;
        }
    }
}
