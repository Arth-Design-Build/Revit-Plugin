﻿using System;
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
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class ElementIntersection : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
           
            try
            {
                Reference pickedObj = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
                if(pickedObj != null)
                {
                    ElementId eleId = pickedObj.ElementId;
                    Element ele = doc.GetElement(eleId);
                    
                    Options gOptions= new Options();
                    gOptions.DetailLevel = ViewDetailLevel.Fine;
                    GeometryElement geom = ele.get_Geometry(gOptions);
                    Solid gSolid = null;

                    foreach(GeometryObject gObj in geom)
                    {
                        GeometryInstance gInst = gObj as GeometryInstance;
                        if(gInst != null)
                        {
                            GeometryElement gEle = gInst.GetInstanceGeometry();
                            foreach (GeometryObject gO in gEle)
                            {
                                gSolid = gO as Solid;
                            }

                        }
                    }

                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    ElementIntersectsSolidFilter filter = new ElementIntersectsSolidFilter(gSolid);
                    ICollection<ElementId> intersects = collector.OfCategory(BuiltInCategory.OST_Roofs).WherePasses(filter).ToElementIds();

                    uidoc.Selection.SetElementIds(intersects);
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
