using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class DetectClashedTag : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Get all tags in the document
            List<Element> tags = new FilteredElementCollector(doc)
                .OfClass(typeof(IndependentTag))
                .ToList();

            //TaskDialog.Show("Tags", tags.Count.ToString());
            // Loop through each tag and check for clashes with other tags
            for (int i = 0; i < tags.Count; i++)
            {
                Element tag = tags[i];
                // Get the location point of the tagged element
                IEnumerable<ElementId> taggedElementIds = ((IndependentTag)tag).GetTaggedLocalElementIds();

                Element taggedElement = doc.GetElement(taggedElementIds.First());
                LocationPoint taggedElementLocation = taggedElement.Location as LocationPoint;
                if (taggedElementLocation == null)
                {
                    continue;
                }

                // Get the bounding box of the tagged element
                BoundingBoxXYZ boundingBox = taggedElement.get_BoundingBox(null);

                // Calculate the topmost point of the bounding box
                XYZ topmostPoint = new XYZ(boundingBox.Min.X, boundingBox.Max.Y, boundingBox.Max.Z);

                // Use the topmost point as the location of the element for clash detection
                XYZ taggedElementPoint = topmostPoint;

                // Loop through all other tags to check for clashes
                for (int j = i + 1; j < tags.Count; j++)
                {
                    Element otherTag = tags[j];
                    // Don't compare a tag to itself
                    if (tag.Id == otherTag.Id)
                    {
                        continue;
                    }

                    // Get the location point of the other tagged element
                    IEnumerable<ElementId> otherTaggedElementIds = ((IndependentTag)otherTag).GetTaggedLocalElementIds();

                    Element otherTaggedElement = doc.GetElement(otherTaggedElementIds.First());
                    LocationPoint otherTaggedElementLocation = otherTaggedElement.Location as LocationPoint;
                    if (otherTaggedElementLocation == null)
                    {
                        continue;
                    }

                    // Get the bounding box of the other tagged element
                    BoundingBoxXYZ otherBoundingBox = otherTaggedElement.get_BoundingBox(null);

                    // Calculate the topmost point of the bounding box
                    XYZ otherTopmostPoint = new XYZ(otherBoundingBox.Min.X, otherBoundingBox.Max.Y, otherBoundingBox.Max.Z);

                    // Use the topmost point as the location of the element for clash detection
                    XYZ otherTaggedElementPoint = otherTopmostPoint;

                    // Calculate the horizontal and vertical distances between the topmost points of the bounding boxes
                    double dx = taggedElementPoint.X - otherTaggedElementPoint.X;
                    double dy = taggedElementPoint.Y - otherTaggedElementPoint.Y;
                    double dz = taggedElementPoint.Z - otherTaggedElementPoint.Z;
                    double distance = Math.Sqrt(dx * dx + dy * dy + dz * dz);

                    // Adjust the threshold values as needed for horizontal and vertical distances
                    double horizontalThreshold = 2;
                    double verticalThreshold = 1;

                    if (distance < horizontalThreshold && Math.Abs(dy) < verticalThreshold && Math.Abs(dz) < verticalThreshold)
                    {
                        // Highlight the clash
                        List<ElementId> idsToHighlight = new List<ElementId>();
                        idsToHighlight.Add(tag.Id);
                        idsToHighlight.Add(otherTag.Id);
                        //uiDoc.Selection.SetElementIds(idsToHighlight);

                        OverrideGraphicSettings ogs = new OverrideGraphicSettings();
                        Color red = new Color(255, 0, 0);
                        Element solidFill = new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement)).Where(q => q.Name.Contains("Solid")).First();

                        ogs.SetProjectionLineColor(red);
                        //ogs.SetProjectionLineWeight(4);

                        using (Transaction t = new Transaction(doc, "Highlight element"))
                        {
                            try
                            {
                                t.Start();
                                uiDoc.ActiveView.SetElementOverrides(idsToHighlight[0], ogs);
                                uiDoc.ActiveView.SetElementOverrides(idsToHighlight[1], ogs);
                                uiDoc.RefreshActiveView();
                                t.Commit();
                            }
                            catch (Exception ex)
                            {
                                TaskDialog.Show("Exception", ex.ToString());
                            }
                        }

                    }
                }
            }

            return Result.Succeeded;
        }
    }
}