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
    public class ViewFilter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            List<ElementId> cats = new List<ElementId>();
            cats.Add(new ElementId(BuiltInCategory.OST_Sections));

            ElementParameterFilter filter = new ElementParameterFilter(ParameterFilterRuleFactory.CreateContainsRule(new ElementId(BuiltInParameter.VIEW_NAME), "WIP", false));

            try
            {
                using (Transaction trans = new Transaction(doc, "Create Plan View"))
                {
                    trans.Start();
                    ParameterFilterElement filterElement = ParameterFilterElement.Create(doc, "My First Filter!", cats, filter);
                    doc.ActiveView.AddFilter(filterElement.Id);
                    doc.ActiveView.SetFilterVisibility(filterElement.Id, false);
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
