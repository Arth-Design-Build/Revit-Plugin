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
                if (!double.TryParse(PromptForOffset(), out offset))
                {
                    TaskDialog.Show("Error", "Invalid offset value");
                    return Result.Failed;
                }

                // Convert the offset to feet (Revit's internal unit)
                offset = UnitUtils.ConvertToInternalUnits(offset / 304.8, UnitTypeId.Feet);

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

        private string PromptForOffset()
        {
            TaskDialog inputDialog = new TaskDialog("Enter Offset")
            {
                MainInstruction = "Enter the offset value in millimeters:",
                MainIcon = TaskDialogIcon.TaskDialogIconNone
            };
            inputDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "OK");
            inputDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Cancel");
            inputDialog.Show();

            return inputDialog.MainInstruction;
        }

        private void AlignSelectedTags(Document doc, IList<Reference> selectedTags, double offset)
        {
            using (Transaction trans = new Transaction(doc, "Align Tags"))
            {
                trans.Start();

                // Sort the tags by their Y-coordinate
                var sortedTags = selectedTags.Select(r => doc.GetElement(r))
                                             .OrderByDescending(e => ((LocationPoint)e.Location).Point.Y)
                                             .ToList();

                // Align the tags with the specified offset
                for (int i = 1; i < sortedTags.Count; i++)
                {
                    Element currentTag = sortedTags[i];
                    Element previousTag = sortedTags[i - 1];

                    XYZ currentTagLocation = ((LocationPoint)currentTag.Location).Point;
                    XYZ previousTagLocation = ((LocationPoint)previousTag.Location).Point;

                    double newY = previousTagLocation.Y - offset;
                    XYZ newLocation = new XYZ(currentTagLocation.X, newY, currentTagLocation.Z);

                    (currentTag.Location as LocationPoint).Move(newLocation - currentTagLocation);
                }

                trans.Commit();
            }
        }
    }
}
