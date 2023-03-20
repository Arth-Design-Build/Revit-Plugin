using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddMEPTag : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Prompt user to upload tag family
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Revit Tag Families (*.rfa)|*.rfa";
            openFileDialog.Title = "Select a Tag Family";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                // User cancelled the dialog
                return Result.Cancelled;
            }
            string tagFamilyPath = openFileDialog.FileName;

            // Load tag family into document
            FamilySymbol tagSymbol = null;
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Load Tag Family");
                Family tagFamily = null;
                bool loaded = doc.LoadFamily(tagFamilyPath, out tagFamily);
                if (!loaded || tagFamily == null)
                {
                    message = "Failed to load tag family";
                    return Result.Failed;
                }
                tagSymbol = doc.GetElement(tagFamily.GetFamilySymbolIds().First()) as FamilySymbol;
                t.Commit();
            }

            // Specify tag mode and orientation
            TagMode tagMode = TagMode.TM_ADDBY_CATEGORY;
            TagOrientation tagOrientation = TagOrientation.Horizontal;

            // Get MEP systems to tag
            List<BuiltInCategory> systemCategories = new List<BuiltInCategory>();
            systemCategories.Add(BuiltInCategory.OST_DuctSystem);
            systemCategories.Add(BuiltInCategory.OST_PipingSystem);
            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(systemCategories);
            IList<Element> systems = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements();

            try
            {
                // Place tags for each MEP system
                using (Transaction trans = new Transaction(doc, "Tag MEP Systems"))
                {
                    trans.Start();
                    foreach (Element system in systems)
                    {
                        Reference systemRef = new Reference(system);
                        LocationPoint systemLocation = system.Location as LocationPoint;
                        XYZ systemPoint = systemLocation.Point;
                        IndependentTag tag = IndependentTag.Create(doc, doc.ActiveView.Id, systemRef, true, tagMode, tagOrientation, systemPoint);
                    }

                    // Commit the transaction
                    trans.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }
        }
    }
}


