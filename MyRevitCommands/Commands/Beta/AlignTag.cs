using Aspose.Cells.Charts;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class AlignTag : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Prompt the user to select tags
                IList<Reference> selectedTags = uiDoc.Selection.PickObjects(ObjectType.Element, "Select tags to align");

                // Prompt the user to enter the offset value
                double offset;

                // Convert the offset to feet (Revit's internal unit)
                offset = UnitUtils.ConvertToInternalUnits(1500 / 304.8, UnitTypeId.Feet);

                // Align the tags
                AlignSelectedTags(doc, selectedTags, offset);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void AlignSelectedTags(Document doc, IList<Reference> selectedTags, double offset)
        {
            using (Transaction trans = new Transaction(doc, "Align Tags"))
            {
                trans.Start();

                // Get the IndependentTag elements from the selection set
                List<IndependentTag> tags = selectedTags
                    .Select(r => doc.GetElement(r))
                    .OfType<IndependentTag>()
                    .Where(e => e.Location != null && e.Location.GetType() == typeof(LocationPoint))
                    .Cast<IndependentTag>()
                    .ToList();

                // Sort the tags by their Y-coordinate
                tags.Sort((t1, t2) =>
                {
                    Location loc1 = t1.Location as Location;
                    Location loc2 = t2.Location as Location;
                    double y1 = loc1 is LocationPoint ? ((LocationPoint)loc1).Point.Y : double.MinValue;
                    double y2 = loc2 is LocationPoint ? ((LocationPoint)loc2).Point.Y : double.MinValue;
                    return y2.CompareTo(y1);
                });

                // Align the tags with the specified offset
                double startY = ((LocationPoint)tags[0].Location).Point.Y;
                foreach (IndependentTag tag in tags)
                {
                    LocationPoint location = tag.Location as LocationPoint;
                    double newY = startY - offset;
                    XYZ newLocation = new XYZ(location.Point.X, newY, location.Point.Z);
                    location.Point = newLocation;
                }

                trans.Commit();
            }
        }
    }
}
