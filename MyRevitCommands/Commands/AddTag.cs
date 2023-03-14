using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using OfficeOpenXml.Packaging.Ionic.Zlib;
using System.Windows.Forms;
using System.Collections.Generic;
using Autodesk.Revit.Creation;
using System.Xml.Linq;

namespace MyRevitCommands
{
    [Transaction(TransactionMode.Manual)]
    public class AddTag : IExternalCommand
    {
        static string PluralSuffix(int n)
        {
            return 1 == n ? "" : "s";
        }

        void TraverseSystems(Autodesk.Revit.DB.Document doc, Autodesk.Revit.DB.View activeView)
        {

            TagMode tmode = TagMode.TM_ADDBY_MATERIAL;
            TagOrientation tOrient = TagOrientation.Horizontal;

            FilteredElementCollector systems
              = new FilteredElementCollector(doc)
                .OfClass(typeof(MEPSystem));

            int i, n;
            string s;
            string[] a;

            StringBuilder message = new StringBuilder();
            using (Transaction trans = new Transaction(doc, "Tag Elements"))
            {
                trans.Start();
                foreach (MEPSystem system in systems)
                {
                    message.AppendLine("System Name: "
                      + system.Name);

                    //message.AppendLine("Base Equipment: " + system.BaseEquipment);

                    Element sysElement = doc.GetElement(system.Id);
                    Reference rf = new Reference(sysElement);
                    LocationPoint loc = sysElement.Location as LocationPoint;
                    XYZ point = loc.Point;
                    IndependentTag newTag = IndependentTag.Create(doc, activeView.Id, rf, true, tmode, tOrient, point);

                    ConnectorSet cs = system.ConnectorManager
                      .Connectors;

                    i = 0;
                    n = cs.Size;
                    a = new string[n];

                    s = string.Format("{0} element{1} in ConnectorManager: ",
                      n, PluralSuffix(n));

                    foreach (Connector c in cs)
                    {
                        Element e = c.Owner;

                        if (null != e)
                        {
                            a[i++] = e.GetType().Name
                              + " " + e.Id.ToString();
                        }
                    }

                    //message.AppendLine(s + string.Join(", ", a));

                    i = 0;
                    n = system.Elements.Size;
                    a = new string[n];

                    s = string.Format(
                      "{0} element{1} in System: ",
                      n, PluralSuffix(n));

                    foreach (Element e in system.Elements)
                    {
                        a[i++] = e.GetType().Name
                          + " " + e.Id.ToString();
                    }

                    //message.AppendLine(s + string.Join(", ", a));
                }
                n = systems.Count<Element>();

                string caption =
                  string.Format("Traverse {0} MEP System{1}",
                  n, (1 == n ? "" : "s"));

                TaskDialog dlg = new TaskDialog(caption);
                dlg.MainContent = message.ToString();
                dlg.Show();
                trans.Commit();
            }
            
        }

        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Autodesk.Revit.DB.View activeView = uidoc.ActiveView;

            TraverseSystems(doc,activeView);
            //Tag Parameters
            TagMode tmode = TagMode.TM_ADDBY_CATEGORY;
            TagOrientation tOrient = TagOrientation.Horizontal;

            FilteredElementCollector systems
              = new FilteredElementCollector(doc)
                .OfClass(typeof(MEPSystem));

            List<BuiltInCategory> cats = new List<BuiltInCategory>();
            cats.Add(BuiltInCategory.OST_Windows);
            cats.Add(BuiltInCategory.OST_Doors);

            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(cats);

            IList<Element> tElements = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements();

            try
            {
                using (Transaction trans = new Transaction(doc, "Tag Elements"))
                {
                    trans.Start();

                    //Tag Elements
                    foreach (Element ele in tElements)
                    {
                        Reference refe = new Reference(ele);
                        //TaskDialog.Show("D and W Ref.", refe.ElementId.ToString());
                        LocationPoint loc = ele.Location as LocationPoint;
                        XYZ point = loc.Point;
                        IndependentTag tag = IndependentTag.Create(doc, doc.ActiveView.Id, refe, true, tmode, tOrient, point);
                    }

                    trans.Commit();
                }
                TaskDialog.Show("Success", "Placed Door and Window Tags");
            }
            catch (Exception e)
            {
                message = e.Message;
                TaskDialog.Show("Failed", "Not Placed Door and Window Tags");
            }

            try
            {
                using (Transaction trans = new Transaction(doc, "Tag Elements"))
                {
                    trans.Start();
                    foreach (MEPSystem sys in systems)
                    {
                        continue;
                    }

                    trans.Commit();
                }
                TaskDialog.Show("Success", "Placed MEP Tags");
            }
            catch (Exception e)
            {
                message = e.Message;
                TaskDialog.Show("Failed", "Not Placed MEP Tags");
            }
            return Result.Succeeded;
        }
    }

}
