using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyRevitCommands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AddDimPipe : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                ReferenceArray pipesRef = new ReferenceArray();
                IList<Pipe> selectedPipes = new List<Pipe>();

                // Get selected pipes
                foreach (ElementId elemId in uidoc.Selection.GetElementIds())
                {
                    Element elem = doc.GetElement(elemId);
                    if (true)
                    {
                        selectedPipes.Add(elem as Pipe);
                        pipesRef.Append(new Reference(elem));
                    }
                }

                // Get bounding box of selected pipes
                BoundingBoxXYZ bbox = GetBoundingBox(selectedPipes);

                // Create dimension lines
                using (Transaction t = new Transaction(doc, "Add dimension lines to pipes"))
                {
                    t.Start();

                    if (bbox != null)
                    {
                        XYZ minPoint = bbox.Min;
                        XYZ maxPoint = bbox.Max;

                        if (minPoint.X == maxPoint.X)
                        {
                            // Vertical pipes
                            XYZ dimLineStart = new XYZ(minPoint.X - 10, minPoint.Y, minPoint.Z);
                            XYZ dimLineEnd = new XYZ(maxPoint.X - 10, maxPoint.Y, maxPoint.Z);
                            Line dimLine = Line.CreateBound(dimLineStart, dimLineEnd);
                            doc.Create.NewDimension(uidoc.ActiveView, dimLine, pipesRef);
                        }
                        else if (minPoint.Y == maxPoint.Y)
                        {
                            // Horizontal pipes
                            XYZ dimLineStart = new XYZ(minPoint.X, minPoint.Y - 10, minPoint.Z);
                            XYZ dimLineEnd = new XYZ(maxPoint.X, maxPoint.Y - 10, maxPoint.Z);
                            Line dimLine = Line.CreateBound(dimLineStart, dimLineEnd);
                            doc.Create.NewDimension(uidoc.ActiveView, dimLine, pipesRef);
                        }
                    }
                    else
                    {
                        TaskDialog.Show("B-Box", "No element to show");
                    }

                    t.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Failed", ex.Message.ToString());
                return Result.Failed;
            }
        }

        private BoundingBoxXYZ GetBoundingBox(IList<Pipe> pipes)
        {
            BoundingBoxXYZ bbox = null;

            foreach (Pipe pipe in pipes)
            {
                BoundingBoxXYZ pipeBbox = pipe.get_BoundingBox(null);
                if (bbox == null)
                {
                    bbox = pipeBbox;
                }
                else
                {
                    XYZ bboxMin = bbox.Min;
                    XYZ bboxMax = bbox.Max;
                    XYZ pipeBboxMin = pipeBbox.Min;
                    XYZ pipeBboxMax = pipeBbox.Max;

                    bboxMin = new XYZ(Math.Min(bboxMin.X, pipeBboxMin.X), Math.Min(bboxMin.Y, pipeBboxMin.Y), Math.Min(bboxMin.Z, pipeBboxMin.Z));
                    bboxMax = new XYZ(Math.Max(bboxMax.X, pipeBboxMax.X), Math.Max(bboxMax.Y, pipeBboxMax.Y), Math.Max(bboxMax.Z, pipeBboxMax.Z));

                    bbox.Min = bboxMin;
                    bbox.Max = bboxMax;
                }
            }

            return bbox;
        }
    }
}