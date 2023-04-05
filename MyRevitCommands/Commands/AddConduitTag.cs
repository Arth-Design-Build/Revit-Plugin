using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddConduitTag: IExternalCommand
    {
        private Document _doc;
        private UIDocument _uiDoc;

        /*
        public AddConduitTag(UIApplication uiapp)
        {
            _uiDoc = uiapp.ActiveUIDocument;
            _doc = _uiDoc.Document;
            _app = uiapp.Application;
        }
        */

        private XYZ MoveRight(XYZ point, double scaleFactor)
        {
            return new XYZ(point.X + scaleFactor, point.Y, point.Z);
        }

        private XYZ MoveDown(XYZ point, double scaleFactor)
        {
            return new XYZ(point.X, point.Y - scaleFactor, point.Z);
        }

        private XYZ MoveLeft(XYZ point, double scaleFactor)
        {
            return new XYZ(point.X - scaleFactor, point.Y, point.Z);
        }

        private XYZ MoveUp(XYZ point, double scaleFactor)
        {
            return new XYZ(point.X, point.Y + scaleFactor, point.Z);
        }

        private IEnumerable<XYZ> Shift(int end, XYZ point, double scaleFactor)
        {
            var moves = new List<System.Func<XYZ, double, XYZ>> { MoveRight, MoveDown, MoveLeft, MoveUp };
            var n = 1;
            var pos = point;
            var timesToMove = 1;

            yield return pos;

            while (true)
            {
                for (var i = 0; i < 2; i++)
                {
                    var move = moves[i % 4];
                    for (var j = 0; j < timesToMove; j++)
                    {
                        if (n >= end) yield break;
                        pos = move(pos, scaleFactor);
                        n++;
                        yield return pos;
                    }
                }
                timesToMove++;
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document _doc = uidoc.Document;
            Selection sel = uidoc.Selection;

            // Prompt the user to make a selection
            var selection = uidoc.Selection;
            IList<Element> selectedElements = sel.PickElementsByRectangle();

            var ConduitFiltered = new List<Element>();
            foreach (var elem in selectedElements)
            {
                if (elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Conduit || elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_ConduitFitting)
                {
                    //FamilyInstance Conduit = (FamilyInstance)elem;
                    ConduitFiltered.Add(elem);
                }
            }

            var scaleFactor = 5.0;

            var avoidLoc = new List<XYZ>();
            foreach (var d in ConduitFiltered)
            {
                BoundingBoxXYZ boundingBox1 = d.get_BoundingBox(_doc.ActiveView);
                XYZ centerPoint = (boundingBox1.Max + boundingBox1.Min) / 2;
                if (centerPoint != null) avoidLoc.Add(new XYZ(centerPoint.X, centerPoint.Y, centerPoint.Z));
            }

            // Get a list of available families for ducts
            FilteredElementCollector fec = new FilteredElementCollector(_doc);
            fec.OfClass(typeof(Family));
            List<Family> families = fec.Cast<Family>().ToList().Where(f => f.Name.Contains("Conduit Tag") || f.Name.Contains("Conduit_Tag")).ToList();
            List<string> familyNames = new List<string>();
            foreach (Family family in families)
            {
                familyNames.Add(family.Name);
            }

            // Create a form with radio buttons to select the family
            System.Windows.Forms.Form forma = new System.Windows.Forms.Form();
            forma.Text = "Select a family";
            forma.StartPosition = FormStartPosition.CenterScreen;
            forma.FormBorderStyle = FormBorderStyle.FixedDialog;
            forma.MinimizeBox = false;
            forma.MaximizeBox = false;
            forma.ShowInTaskbar = false;
            forma.AutoScroll = true;
            forma.ClientSize = new Size(500, 200);
            forma.BackColor = System.Drawing.Color.LightGray;

            GroupBox groupBox = new GroupBox();
            groupBox.Location = new System.Drawing.Point(10, 10);
            groupBox.Size = new Size(480, 140);
            forma.Controls.Add(groupBox);

            int y = 20;
            foreach (Family family in families)
            {
                RadioButton radioButton = new RadioButton();
                radioButton.Text = family.Name;
                radioButton.Location = new System.Drawing.Point(20, y);
                radioButton.AutoSize = true;
                radioButton.Tag = family;
                groupBox.Controls.Add(radioButton);
                y += 25;
            }

            System.Windows.Forms.Button okButton = new System.Windows.Forms.Button();
            okButton.Text = "OK";
            okButton.DialogResult = DialogResult.OK;
            okButton.Location = new System.Drawing.Point(180, 160);
            forma.Controls.Add(okButton);

            System.Windows.Forms.Button cancelButton = new System.Windows.Forms.Button();
            cancelButton.Text = "Cancel";
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Location = new System.Drawing.Point(260, 160);
            forma.Controls.Add(cancelButton);

            // Show the form and get the selected family
            Family selectedFamily = null;
            if (forma.ShowDialog() == DialogResult.OK)
            {
                foreach (System.Windows.Forms.Control control in groupBox.Controls)
                {
                    if (control is RadioButton && ((RadioButton)control).Checked)
                    {
                        selectedFamily = (Family)((RadioButton)control).Tag;
                        break;
                    }
                }
            }

            var tagFamilyName = "Conduit Tag"; // replace with your desired tag family name
            var tagFamily = new FilteredElementCollector(_doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .FirstOrDefault(f => f.Name.Contains("Conduit Tag") || f.Name.Contains("Conduit_Tag"));

            var tagSymbols = new FilteredElementCollector(_doc)
            .OfCategory(BuiltInCategory.OST_ConduitTags)
            .OfClass(typeof(FamilySymbol))
            .Cast<FamilySymbol>()
            .FirstOrDefault(fs => fs.Family.Id == tagFamily.Id);

            if (tagSymbols == null)
            {
                //TaskDialog.Show("Error", "No Tag Symbols Found");
                return Result.Failed;
            }

            //TaskDialog.Show("Count", ConduitFiltered.Count.ToString());

            foreach (var d in ConduitFiltered)
            {
                BoundingBoxXYZ boundingBox2 = d.get_BoundingBox(_doc.ActiveView);
                XYZ centerPoint = (boundingBox2.Max + boundingBox2.Min) / 2;
                if (centerPoint == null)
                {
                    //TaskDialog.Show("CenterPoint", "Null");
                    continue;
                }
                var levelPoint = new XYZ(centerPoint.X, centerPoint.Y, centerPoint.Z);
                var R = new Reference(d);
                var tx = new Transaction(_doc);
                tx.Start("Tag Conduits");
                
                var IT = IndependentTag.Create(_doc, uidoc.ActiveView.Id, R, true, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, levelPoint);
                IT.ChangeTypeId(tagSymbols.Id);
                tx.Commit();

                var tagBB = IT.get_BoundingBox((Autodesk.Revit.DB.View)_doc.GetElement(IT.OwnerViewId));
                var globalMax = tagBB.Max;
                var globalMin = tagBB.Min;
                var BBloc = new XYZ((globalMax.X), (globalMax.Y + globalMin.Y) / 2, globalMax.Z);
                var avoidPoints = Shift(50, BBloc, scaleFactor);

                foreach (var point in avoidPoints)
                {
                    var closestConduit = FindClosestConduit(point, avoidLoc);
                    if (closestConduit != null)
                    {
                        var distance = closestConduit.DistanceTo(point);
                        if (distance < scaleFactor)
                        {
                            var direction = (point - closestConduit).Normalize();
                            var newPoint = closestConduit + (direction * scaleFactor);
                            var tx1 = new Transaction(_doc);
                            tx1.Start("Modification of Tags");
                            IT.TagHeadPosition = newPoint;
                            IT.LeaderEndCondition = LeaderEndCondition.Free;
                            tx1.Commit();
                            break;
                        }
                    }
                }
            }
            return Result.Succeeded;
        }

        private XYZ FindClosestConduit(XYZ point, List<XYZ> avoidLoc)
        {
            XYZ closestConduit = null;
            var closestDistance = double.MaxValue;
            foreach (var Conduit in avoidLoc)
            {
                var distance = Conduit.DistanceTo(point);
                if (distance < closestDistance)
                {
                    closestConduit = Conduit;
                    closestDistance = distance;
                }
            }
            return closestConduit;
        }
    }
}
