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
    public class PlaceLineElement : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Level level = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsElementType()
                .Cast<Level>()
                .First(x => x.Name == "Ground Floor");
            XYZ p1 = new XYZ(-10,-10,0);
            XYZ p2 = new XYZ(10,-10,0);
            XYZ p3 = new XYZ(15,0,0);
            XYZ p4 = new XYZ(10,10,0);
            XYZ p5 = new XYZ(-10,10,0);

            List<Curve> curves = new List<Curve>();
            Line l1 = Line.CreateUnbound(p1,p2);
            Arc l2 = Arc.Create(p2, p4, p3);
            Line l3 = Line.CreateUnbound(p4,p5);
            Line l4 = Line.CreateUnbound(p5,p1);

            curves.Add(l1);
            curves.Add(l2);
            curves.Add(l3);
            curves.Add(l4);

            try
            {
                using (Transaction trans = new Transaction(doc, "Place Line Element"))
                {
                    trans.Start();
                    foreach(Curve c in curves)
                    {
                        Wall.Create(doc, c, level.Id, false);
                    }
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
