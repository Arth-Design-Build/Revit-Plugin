using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class PlanView : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            ViewFamilyType viewFamily = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .First(x => x.ViewFamily == ViewFamily.FloorPlan);
            Level level = new FilteredElementCollector(doc)
               .OfCategory(BuiltInCategory.OST_Levels)
               .WhereElementIsElementType()
               .Cast<Level>()
               .First(x => x.Name == "Ground Floor");

            try
            {
                using (Transaction trans = new Transaction(doc, "Create Plan View"))
                {
                    trans.Start();
                    ViewPlan vPlan = ViewPlan.Create(doc, viewFamily.Id, level.Id);
                    vPlan.Name = "Our first plan!";

                    trans.Commit();
                }
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                message= e.Message;
                return Result.Failed;
            }
            
        }
    }
}
