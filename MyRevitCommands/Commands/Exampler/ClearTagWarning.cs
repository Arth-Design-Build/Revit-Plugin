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
    public class ClearTagWarning : IExternalCommand
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

            // Clear the tag highlights
            OverrideGraphicSettings clearSettings = new OverrideGraphicSettings();
            clearSettings.SetProjectionLineColor(new Color(0, 0, 0));
            clearSettings.SetProjectionLineWeight(1);

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Clear Tag Highlights");
                foreach (Element tag in tags)
                {
                    tag.Document.ActiveView.SetElementOverrides(tag.Id, clearSettings);
                }
                tx.Commit();
            }

            return Result.Succeeded;
        }
    }
}
